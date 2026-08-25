# -*- coding: UTF-8 -*-
"""Build Comment Commander into an installable .nvda-addon package.

Uses nothing but the Python standard library, so it runs with any Python 3.8+
without installing SCons or the NVDA add-on template.

	python build.py

Produces commentCommander-<version>.nvda-addon in this directory.
"""

import array
import os
import re
import shutil
import struct
import sys
import zipfile

from buildVars import addon_info, excludedFiles

HERE = os.path.dirname(os.path.abspath(__file__))
ADDON_DIR = os.path.join(HERE, "addon")

#: Never ship these.
EXCLUDED_NAMES = {"__pycache__", ".git", ".svn", "desktop.ini", "Thumbs.db"}
EXCLUDED_SUFFIXES = (".pyc", ".pyo", ".po", ".pot", ".bak")


# ---------------------------------------------------------------------------
# .po -> .mo compilation (a trimmed down port of Python's Tools/i18n/msgfmt.py)
# ---------------------------------------------------------------------------


class _Entry:
	"""One accumulating .po entry."""

	def __init__(self):
		self.ctxt = ""
		self.msgid = ""
		self.msgidPlural = ""
		self.msgstr = ""
		self.plurals = {}
		self.isPlural = False
		self.fuzzy = False

	def key(self):
		key = self.msgid + "\x00" + self.msgidPlural if self.isPlural else self.msgid
		return self.ctxt + "\x04" + key if self.ctxt else key

	def value(self):
		if self.isPlural:
			return "\x00".join(self.plurals[i] for i in sorted(self.plurals))
		return self.msgstr

	def isTranslated(self):
		# Fuzzy entries are unreviewed guesses, and empty ones fall back to the
		# original English, so neither belongs in the compiled catalogue.
		if self.fuzzy:
			return False
		if self.isPlural:
			return any(self.plurals.values())
		return bool(self.msgstr)


def _parsePo(path):
	"""Return {key: translation} for a .po file, plural forms included."""
	messages = {}
	entry = _Entry()
	section = None
	pluralIndex = 0

	def store():
		nonlocal entry
		# The empty msgid holds the catalogue header, which gettext wants kept.
		if entry.isTranslated() or (entry.msgid == "" and entry.msgstr and not entry.ctxt):
			messages[entry.key()] = entry.value()
		entry = _Entry()

	with open(path, encoding="utf-8") as f:
		lines = f.readlines()

	for line in lines:
		line = line.rstrip("\n")
		stripped = line.strip()
		if stripped.startswith("#"):
			if stripped.startswith("#,") and "fuzzy" in stripped:
				entry.fuzzy = True
			continue
		if not stripped:
			continue
		if line.startswith("msgctxt "):
			if section in ("str", "str_plural"):
				store()
			section = "ctxt"
			entry.ctxt = _unquote(line[8:])
		elif line.startswith("msgid_plural "):
			section = "id_plural"
			entry.isPlural = True
			entry.msgidPlural = _unquote(line[13:])
		elif line.startswith("msgid "):
			if section in ("str", "str_plural"):
				store()
			section = "id"
			entry.msgid = _unquote(line[6:])
		elif line.startswith("msgstr["):
			match = re.match(r'msgstr\[(\d+)\]\s*(.*)', line)
			section = "str_plural"
			pluralIndex = int(match.group(1))
			entry.plurals[pluralIndex] = _unquote(match.group(2))
		elif line.startswith("msgstr "):
			section = "str"
			entry.msgstr = _unquote(line[7:])
		else:
			# A bare quoted string continues whichever field came last.
			text = _unquote(line)
			if section == "ctxt":
				entry.ctxt += text
			elif section == "id":
				entry.msgid += text
			elif section == "id_plural":
				entry.msgidPlural += text
			elif section == "str":
				entry.msgstr += text
			elif section == "str_plural":
				entry.plurals[pluralIndex] += text
	store()
	return messages


def _unquote(text):
	text = text.strip()
	if not text.startswith('"'):
		return ""
	# Trust the .po escaping rules; they match Python's.
	return text[1:-1].encode().decode("unicode_escape").encode("latin-1").decode("utf-8")


def compilePo(poPath, moPath):
	"""Compile a .po file into the binary .mo format gettext expects."""
	messages = _parsePo(poPath)
	keys = sorted(messages.keys())
	offsets = []
	ids = strs = b""
	for key in keys:
		encodedId = key.encode("utf-8")
		encodedStr = messages[key].encode("utf-8")
		offsets.append((len(ids), len(encodedId), len(strs), len(encodedStr)))
		ids += encodedId + b"\x00"
		strs += encodedStr + b"\x00"

	keyCount = len(keys)
	keyStart = 7 * 4 + keyCount * 16
	valueStart = keyStart + len(ids)
	keyOffsets = []
	valueOffsets = []
	for o1, l1, o2, l2 in offsets:
		keyOffsets += [l1, o1 + keyStart]
		valueOffsets += [l2, o2 + valueStart]

	output = struct.pack(
		"Iiiiiii",
		0x950412DE,  # gettext magic
		0,  # version
		keyCount,
		7 * 4,  # offset of the key table
		7 * 4 + keyCount * 8,  # offset of the value table
		0,
		0,
	)
	output += array.array("i", keyOffsets + valueOffsets).tobytes()
	output += ids + strs

	os.makedirs(os.path.dirname(moPath), exist_ok=True)
	with open(moPath, "wb") as f:
		f.write(output)
	return keyCount


def buildTranslations():
	"""Compile every locale/<lang>/LC_MESSAGES/nvda.po into a .mo next to it."""
	localeDir = os.path.join(ADDON_DIR, "locale")
	if not os.path.isdir(localeDir):
		return
	for language in sorted(os.listdir(localeDir)):
		poPath = os.path.join(localeDir, language, "LC_MESSAGES", "nvda.po")
		if not os.path.isfile(poPath):
			continue
		moPath = os.path.join(localeDir, language, "LC_MESSAGES", "nvda.mo")
		count = compilePo(poPath, moPath)
		print(f"  translated {language}: {count} strings")


# ---------------------------------------------------------------------------
# Manifest
# ---------------------------------------------------------------------------


def writeManifest():
	path = os.path.join(ADDON_DIR, "manifest.ini")
	info = addon_info
	lines = [
		"name = {}".format(info["addon_name"]),
		'summary = "{}"'.format(info["addon_summary"]),
		'description = """{}"""'.format(info["addon_description"]),
		'author = "{}"'.format(info["addon_author"]),
		"version = {}".format(info["addon_version"]),
		"minimumNVDAVersion = {}".format(info["addon_minimumNVDAVersion"]),
		"lastTestedNVDAVersion = {}".format(info["addon_lastTestedNVDAVersion"]),
	]
	if info.get("addon_url"):
		lines.append("url = {}".format(info["addon_url"]))
	if info.get("addon_sourceURL"):
		lines.append("sourceURL = {}".format(info["addon_sourceURL"]))
	if info.get("addon_docFileName"):
		lines.append("docFileName = {}".format(info["addon_docFileName"]))
	if info.get("addon_license"):
		lines.append("license = {}".format(info["addon_license"]))
	if info.get("addon_licenseURL"):
		lines.append("licenseURL = {}".format(info["addon_licenseURL"]))
	if info.get("addon_updateChannel"):
		lines.append("updateChannel = {}".format(info["addon_updateChannel"]))
	with open(path, "w", encoding="utf-8") as f:
		f.write("\n".join(lines) + "\n")
	print(f"  manifest: {path}")


# ---------------------------------------------------------------------------
# Packaging
# ---------------------------------------------------------------------------


def shouldInclude(path, name):
	if name in EXCLUDED_NAMES:
		return False
	if name.endswith(EXCLUDED_SUFFIXES):
		return False
	relative = os.path.relpath(path, ADDON_DIR).replace(os.sep, "/")
	return relative not in excludedFiles


def buildPackage():
	name = addon_info["addon_name"]
	version = addon_info["addon_version"]
	target = os.path.join(HERE, f"{name}-{version}.nvda-addon")
	if os.path.exists(target):
		os.remove(target)
	fileCount = 0
	with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
		for root, dirs, files in os.walk(ADDON_DIR):
			dirs[:] = [d for d in dirs if shouldInclude(os.path.join(root, d), d)]
			for fileName in sorted(files):
				fullPath = os.path.join(root, fileName)
				if not shouldInclude(fullPath, fileName):
					continue
				archiveName = os.path.relpath(fullPath, ADDON_DIR).replace(os.sep, "/")
				archive.write(fullPath, archiveName)
				fileCount += 1
	print(f"  packaged {fileCount} files")
	return target


def clean():
	for root, dirs, files in os.walk(ADDON_DIR):
		for d in list(dirs):
			if d == "__pycache__":
				shutil.rmtree(os.path.join(root, d), ignore_errors=True)
				dirs.remove(d)
	for fileName in os.listdir(HERE):
		if fileName.endswith(".nvda-addon"):
			os.remove(os.path.join(HERE, fileName))
	print("Cleaned.")


def main():
	if len(sys.argv) > 1 and sys.argv[1] == "clean":
		clean()
		return 0
	print("Building Comment Commander...")
	writeManifest()
	buildTranslations()
	target = buildPackage()
	size = os.path.getsize(target)
	print(f"\nDone: {target} ({size:,} bytes)")
	print("Install it by pressing Enter on the file, or via NVDA's Add-on Store.")
	return 0


if __name__ == "__main__":
	sys.exit(main())
