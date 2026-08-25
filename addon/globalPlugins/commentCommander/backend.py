# -*- coding: UTF-8 -*-
# Comment Commander: chooses the right Office backend for the current focus.
# Copyright (C) 2026 Sensotec
# This file is covered by the GNU General Public License version 2.

"""Backend selection.

The two backends expose the same functions, so everything above this module
works against whichever one the focused application calls for.
"""

import addonHandler

from . import excelAccess, wordAccess

addonHandler.initTranslation()

#: Backends in the order they get asked. Each answers isFocused() cheaply.
BACKENDS = (wordAccess, excelAccess)

#: Identifies a backend without importing it, for comparisons in the UI.
WORD = "word"
EXCEL = "excel"


def kindOf(module) -> str:
	return EXCEL if module is excelAccess else WORD


def _isFocused(module) -> bool:
	if module is wordAccess:
		return module.isWordFocused()
	return module.isExcelFocused()


def isOfficeFocused() -> bool:
	"""True when any supported Office application has the focus.

	COM-free, because this runs on every keypress matching one of the add-on's
	gestures.
	"""
	return any(_isFocused(module) for module in BACKENDS)


def isWordFocused() -> bool:
	return wordAccess.isWordFocused()


def getBackend():
	"""Return (module, document) for the focused application, or (None, None).

	Both backends refuse to look past the focused window, so this can never
	reach a document belonging to an application running in the background.
	"""
	for module in BACKENDS:
		if not _isFocused(module):
			continue
		document = module.getDocumentObject()
		if document is not None:
			return module, document
	return None, None


def supportsRevisions(module) -> bool:
	"""Only Word has tracked changes of the kind this add-on lists."""
	return module is wordAccess


def supportsAddComment(module) -> bool:
	return True
