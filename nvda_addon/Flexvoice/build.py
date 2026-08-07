#!/usr/bin/env python3
"""Package the FlexVoice NVDA addon as a .nvda-addon file (ZIP archive)."""

import os
import sys
import zipfile

if sys.version_info < (3, 0):
	raise Exception("Python 3 required")

ADDON_DIR = os.path.dirname(os.path.abspath(__file__))
SYNTH_DIR = os.path.join(ADDON_DIR, "synthDrivers")
OUTPUT_NAME = "flexvoice.nvda-addon"
OUTPUT_PATH = os.path.join(ADDON_DIR, OUTPUT_NAME)

# Explicit files to include
ADDON_FILES = {
	# Root addon files
	"manifest.ini": os.path.join(ADDON_DIR, "manifest.ini"),

	# SynthDriver Python files
	"synthDrivers/flexvoice.py": os.path.join(SYNTH_DIR, "flexvoice.py"),
	"synthDrivers/_flexvoice.py": os.path.join(SYNTH_DIR, "_flexvoice.py"),
	"synthDrivers/_ipc.py": os.path.join(SYNTH_DIR, "_ipc.py"),

	# Wrapper DLL
	"synthDrivers/fvwrap.dll": os.path.join(SYNTH_DIR, "fvwrap.dll"),

	# Engine DLLs
	"synthDrivers/FlexVoice_3_01_001.dll": os.path.join(SYNTH_DIR, "FlexVoice_3_01_001.dll"),
	"synthDrivers/FVFarmServer_3_01_001.dll": os.path.join(SYNTH_DIR, "FVFarmServer_3_01_001.dll"),
	"synthDrivers/FVNetClient.dll": os.path.join(SYNTH_DIR, "FVNetClient.dll"),
	"synthDrivers/FVNetClient_VBFile.dll": os.path.join(SYNTH_DIR, "FVNetClient_VBFile.dll"),
	"synthDrivers/FVNetServer.dll": os.path.join(SYNTH_DIR, "FVNetServer.dll"),
}

# Voice data directories, one per language. Each must sit directly under
# synthDrivers/ so EngineFactory can be pointed at synthDrivers/ as its data
# root and load whichever language is selected.
LANGUAGE_DIRS = ("English", "Hungarian")


def collect_language_files():
	"""Recursively collect every file from each shipped language directory."""
	files = {}
	for dirName in LANGUAGE_DIRS:
		langDir = os.path.join(SYNTH_DIR, dirName)
		if not os.path.isdir(langDir):
			print(f"  note: {dirName}/ not present, skipping")
			continue
		for root, dirs, filenames in os.walk(langDir):
			for fn in filenames:
				full_path = os.path.join(root, fn)
				rel = os.path.relpath(full_path, SYNTH_DIR)
				arc_name = "synthDrivers/" + rel.replace("\\", "/")
				files[arc_name] = full_path
	return files


def main():
	all_files = dict(ADDON_FILES)
	all_files.update(collect_language_files())

	missing = []
	for arc_name, src_path in all_files.items():
		if not os.path.exists(src_path):
			missing.append((arc_name, src_path))

	if missing:
		print("WARNING: Missing files:")
		for arc_name, src_path in missing:
			print(f"  {arc_name} -> {src_path}")
		print()

	# Only require the Python files and manifest to exist
	required = ["manifest.ini", "synthDrivers/flexvoice.py",
				"synthDrivers/_flexvoice.py"]
	for r in required:
		if r in [m[0] for m in missing]:
			print(f"ERROR: Required file missing: {r}")
			sys.exit(1)

	print(f"Creating {OUTPUT_NAME}...")
	count = 0
	with zipfile.ZipFile(OUTPUT_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
		for arc_name, src_path in sorted(all_files.items()):
			if os.path.exists(src_path):
				zf.write(src_path, arc_name)
				size = os.path.getsize(src_path)
				print(f"  + {arc_name} ({size:,} bytes)")
				count += 1
			else:
				print(f"  - {arc_name} (SKIPPED - not found)")

	print(f"\nCreated {OUTPUT_PATH}")
	print(f"Size: {os.path.getsize(OUTPUT_PATH):,} bytes")
	print(f"Files: {count}")


if __name__ == "__main__":
	main()
