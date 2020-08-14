#!/opt/local/bin/python3.7
"""
Usage:
  c1-mass-relink.py   (--selected | --collection | --all) --new-location <path> [options]

Options:
  --selected         Change only images currently selected in Capture One
  --collection       Change the entire collection in the currently opened in Capture One catalogue
  --progress         Show TQDM progress bar instead of verbose text output
  --progress-gui     Show TQDM GUI progress bar in addition to verbose text output
  --verbose          Output info on relinked images
  --dry-run          Search for images but don't actually relink them
"""
import sys
import os
import docopt
import pathlib
import functools
import re
import ctypes

import appscript
import tqdm

paths_ignore  = re.compile(r"(.*(\.apdisk|\.DS_Store|\.VolumeIcon.icns|\.plist|\.xml|\.[xX][mM][pP])$|.*(\.Spotlight-V100/|\.fseventsd/|\.mpaur2).*)")

class PhotoNameSizeKey(object):
	def __init__(self, path, size = None):
		self.path = path
		self.name = path.name
		self.size = None
		if size:
			self.size = size
		else:
			self.size = self.path.stat().st_size

	@property
	def name_size(self):
		return (self.name + "," + str(self.size))

	# Implement __hash__ and __eq__ so this can serve as a key for a dictionary
	def __hash__(self):
		return(hash((self.name, self.size)))
	
	def __eq__(self, other):
		return isinstance(other, self.__class__) and self.name==other.name and self.size==other.size
			

def generate_directory_dict(path):
	ret = { }
	files = path.rglob("*")
	for f in files:
		if f.is_file() and not paths_ignore.match(f.as_posix()):
			key = PhotoNameSizeKey(f)
			if key not in ret:
				ret[key] = f
			else:
				raise ValueError("Files {} and {} found with same size and name".format(f.as_posix(), ret[key].as_posix()))
	return ret

class C1Image(object):
	def __init__(self, c1_image_object):
		self._c1_image_object = c1_image_object

	@property
	@functools.lru_cache()
	def path(self):
		""" The full path to original image file (R/O) """
		return pathlib.Path(self._c1_image_object.path.get(timeout=3600))

	@property
	@functools.lru_cache()
	def filesize(self):
		""" The image's file size as per Capture Ones database (R/O) """
		# time for a rant: we get what is essential a signed 32 bit integer from C1 via
		# the apple events bridge. Considering that other applications happily give us
		# a unsigned 32 or 64 bit integer, it seems C1 uses the int32 internally to represent
		# filesizes. While images >2G and <4G do not seem to cause problems in C1, this seems
		# more by chance than by design. Need to work around this sloopyness, so we do the
		# conversion via python ctypes
		c1_filesize = self._c1_image_object.file_size.get(timeout=3600)
		return ctypes.c_uint32(ctypes.c_int32(c1_filesize).value).value
	
	@property
	@functools.lru_cache()
	def name(self):
		""" The name of the image file (R/O) """
		return self._c1_image_object.name.get(timeout=3600)
	
	@property
	@functools.lru_cache()
	def id(self):
		""" The unique identifier of the image (R/O) """
		return self._c1_image_object.id.get(timeout=3600)
	
	@functools.lru_cache()
	def photo_name_size_key(self):
		return(PhotoNameSizeKey(self.path, size=int(self.filesize)))
	
	def relink(self,new_path):
		self._c1_image_object.relink(to_path=new_path.as_posix(), waitreply=True)
		

# parse cmd options
arguments = docopt.docopt(__doc__)

if isinstance(arguments["<path>"], str):
	arguments["<path>"] = pathlib.Path(arguments["<path>"])

# Walk the new location and build a file list dictionary
new_location_files = generate_directory_dict(arguments["<path>"])




CaptureOne = appscript.app("Capture One 20")

images = [ ]
if arguments["--all"]:
	images = CaptureOne.current_document.images.get(timeout=1800)
elif arguments["--collection"]:
	images = CaptureOne.current_document.current_collection.images()
elif arguments["--selected"]:
	variants = CaptureOne.selected_variants()
	for variant in variants:
		images.append(variant.parent_image.get())
else:
	raise ValueError("Don't know what variants/images to use, please specify --all, --collection or --selected")


if arguments["--progress"]:
	image_iterator = tqdm.tqdm(images, unit="Image", unit_scale=False, leave=True, position=0)
elif arguments["--progress-gui"]:
	image_iterator = tqdm.tqdm_gui(images, unit="Image", unit_scale=False)
else:
	image_iterator = images


for img_ae_obj in image_iterator:
	image = C1Image(img_ae_obj)

	matched_image = None
	if image.photo_name_size_key() in new_location_files:
		matched_image = new_location_files[image.photo_name_size_key()]

	log = None
	if arguments["--progress"]:
		log = tqdm.tqdm.write
	else:
		log = print

	log_message = None
	log_nonverbose = False
	if image.path.exists():
		log_message = "\tSKIPPING - relinking not necessary"
	elif matched_image:
		log_message = "\tRELINKED to \"{}\"".format(matched_image.as_posix())
	else:
		log_nonverbose = True
		log_message = "\tNO MATCHING IMAGE FOUND"
		log_message+= " (size {} Bytes)".format(image.filesize)
	if arguments["--verbose"] or log_nonverbose == True:
		log("{} (ID {}){}".format(image.name, image.id,log_message))
	
	# Image in C1 Catalog is not connected, we found a matching image file and are not asked for a dry-run...
	if not image.path.exists() and matched_image and not arguments["--dry-run"]:
		image.relink(matched_image)
