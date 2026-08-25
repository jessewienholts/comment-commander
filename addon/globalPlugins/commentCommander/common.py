# -*- coding: UTF-8 -*-
# Comment Commander: pieces shared by the Microsoft Office backends.
# Copyright (C) 2026 Sensotec
# This file is covered by the GNU General Public License version 2.

"""Data classes and COM plumbing shared by the Word and Excel backends.

Everything the user interface sees is defined here, so the dialogs never need
to know which Office application the items came from.
"""

import datetime
from dataclasses import dataclass, field
from typing import Any, Optional

import addonHandler
import api
import comtypes.automation
import comtypes.client
import comtypes.client.dynamic
import oleacc
import winUser
from comtypes import COMError

addonHandler.initTranslation()


class OfficeAccessError(Exception):
	"""Raised when an Office application could not be reached or refused an operation."""


#: Kept under the old name so existing callers keep working.
WordAccessError = OfficeAccessError

#: Values of CommentItem.kind. Word comments carry KIND_NONE.
KIND_NONE = ""
KIND_NOTE = "note"
KIND_THREADED = "comment"

#: Human readable labels for the wdRevisionType enumeration.
REVISION_TYPE_LABELS = {
	# Translators: A type of tracked change in Microsoft Word.
	1: _("Insertion"),
	# Translators: A type of tracked change in Microsoft Word.
	2: _("Deletion"),
	# Translators: A type of tracked change in Microsoft Word.
	3: _("Formatting"),
	# Translators: A type of tracked change in Microsoft Word.
	4: _("Paragraph number"),
	# Translators: A type of tracked change in Microsoft Word.
	5: _("Displayed field"),
	# Translators: A type of tracked change in Microsoft Word.
	6: _("Reconcile"),
	# Translators: A type of tracked change in Microsoft Word.
	7: _("Conflict"),
	# Translators: A type of tracked change in Microsoft Word.
	8: _("Style"),
	# Translators: A type of tracked change in Microsoft Word.
	9: _("Replacement"),
	# Translators: A type of tracked change in Microsoft Word.
	10: _("Paragraph property"),
	# Translators: A type of tracked change in Microsoft Word.
	11: _("Table property"),
	# Translators: A type of tracked change in Microsoft Word.
	12: _("Section property"),
	# Translators: A type of tracked change in Microsoft Word.
	13: _("Style definition"),
	# Translators: A type of tracked change in Microsoft Word.
	14: _("Moved from"),
	# Translators: A type of tracked change in Microsoft Word.
	15: _("Moved to"),
	# Translators: A type of tracked change in Microsoft Word.
	16: _("Cell insertion"),
	# Translators: A type of tracked change in Microsoft Word.
	17: _("Cell deletion"),
	# Translators: A type of tracked change in Microsoft Word.
	18: _("Cell merge"),
}


def normalize(text: Optional[str], maxLength: Optional[int] = None) -> str:
	"""Flatten an Office string into something a list control can show."""
	if not text:
		return ""
	# Word terminates ranges with \r and uses \x07 for cell and row marks.
	for ch in ("\r", "\n", "\v", "\x07", "\x0c"):
		text = text.replace(ch, " ")
	text = " ".join(text.split())
	if maxLength is not None and len(text) > maxLength:
		text = text[: maxLength - 1] + "…"
	return text


def safeGet(getter, default=None):
	"""Read one COM property, tolerating versions that do not expose it."""
	try:
		value = getter()
	except (COMError, AttributeError, NameError, TypeError, ValueError):
		return default
	return default if value is None else value


def readValue(getter, default=None):
	"""Read a value, refusing to invoke it even if COM hands back a callable.

	This exists because of a genuine hazard. Excel exposes the text of both
	annotation kinds as ``Text(Text, Start, Overwrite)``, whose arguments are
	all optional, and Microsoft documents that calling it with no arguments
	*deletes* the existing text. Reading through a property get is what VBA
	does and is safe; calling it is not. So if comtypes hands back a bound
	method here, we return the default rather than risk wiping a comment.
	"""
	try:
		value = getter()
	except (COMError, AttributeError, NameError, TypeError, ValueError):
		return default
	if value is None or callable(value):
		return default
	return value


def flagAsMethod(com, *names) -> None:
	"""Tell comtypes these names are methods, never properties.

	Without this, comtypes first tries a property get, which for a COM method
	whose arguments are all optional actually *runs* it. Calling
	``cell.AddCommentThreaded`` unflagged would therefore create an empty
	comment as a side effect of merely looking the name up.
	"""
	try:
		com._FlagAsMethod(*names)
	except Exception:
		pass


def shareDispatchIds(template, target) -> None:
	"""Copy a resolved dispatch-ID cache from one COM object onto another.

	comtypes' dynamic dispatch caches dispatch IDs per wrapper instance, and
	Office hands out a brand new wrapper for every item in a collection.
	Without this, reading a property costs a GetIDsOfNames round trip on top of
	the Invoke, doubling the COM traffic for the whole document. Items of one
	collection share a COM type, so their IDs are interchangeable.

	Deliberately defensive: if comtypes ever stops exposing this, we silently
	fall back to the slower path rather than breaking.
	"""
	if template is None or target is None or template is target:
		return
	try:
		source = template.__dict__.get("_ids")
		destination = target.__dict__.get("_ids")
		if isinstance(source, dict) and isinstance(destination, dict) and not destination:
			destination.update(source)
	except Exception:
		pass


class Capabilities:
	"""Which optional properties this document actually answers.

	Older Office versions raise on newer members. Retrying on every item would
	mean an exception crossing the COM boundary hundreds of times, so we find
	out once and then stop asking.
	"""

	def __init__(self):
		self._failed = set()

	def read(self, flag: str, getter, default=None):
		if flag in self._failed:
			return default
		try:
			value = getter()
		except (COMError, AttributeError, NameError, TypeError, ValueError):
			self._failed.add(flag)
			return default
		return default if value is None else value


def dispatchFromWindow(hwnd: int):
	"""Fetch an Office object model handle from a window handle, or None."""
	try:
		pDispatch = oleacc.AccessibleObjectFromWindow(
			hwnd,
			winUser.OBJID_NATIVEOM,
			interface=comtypes.automation.IDispatch,
		)
	except (COMError, WindowsError):
		return None
	if not pDispatch:
		return None
	return comtypes.client.dynamic.Dispatch(pDispatch)


def hasWindowClassInAncestry(obj, classNames) -> bool:
	"""True when one of classNames sits in the focused window's ancestry."""
	hwnd = getattr(obj, "windowHandle", 0)
	depth = 0
	while hwnd and depth < 12:
		try:
			if winUser.getClassName(hwnd) in classNames:
				return True
		except Exception:
			pass
		hwnd = winUser.getAncestor(hwnd, winUser.GA_PARENT)
		depth += 1
	return False


def getFocusAppName() -> str:
	"""The NVDA application name of whatever has the focus, or an empty string."""
	try:
		focus = api.getFocusObject()
	except Exception:
		return ""
	if focus is None:
		return ""
	appModule = getattr(focus, "appModule", None)
	return getattr(appModule, "appName", "") if appModule is not None else ""


@dataclass
class CommentItem:
	"""One comment, or one reply within a thread.

	Covers Word comments as well as both kinds of Excel annotation. Fields that
	do not apply to a given source simply stay at their default.
	"""

	index: int
	author: str = ""
	date: Optional[datetime.datetime] = None
	text: str = ""
	#: Word: the annotated text. Excel: the contents of the commented cell.
	anchorText: str = ""
	#: Word only, filled in lazily because it forces Word to paginate.
	page: Optional[int] = None
	#: Word only. Excel does not expose the resolved state at all.
	done: Optional[bool] = None
	parentIndex: Optional[int] = None
	replyCount: int = 0
	threadPosition: int = 0
	comObject: Any = field(default=None, repr=False)
	#: Cached Word Range or Excel cell, so we never fetch it from Office twice.
	scopeObj: Any = field(default=None, repr=False)
	#: Word: character offsets of the anchor.
	startPos: Optional[int] = None
	endPos: Optional[int] = None
	#: False while only the position fields have been read.
	detailsLoaded: bool = False
	#: Excel: which kind of annotation this is.
	kind: str = KIND_NONE
	#: Excel: where the annotation lives.
	sheetName: str = ""
	cellAddress: str = ""
	sheetIndex: int = 0
	row: int = 0
	column: int = 0

	@property
	def isReply(self) -> bool:
		return self.parentIndex is not None

	@property
	def statusText(self) -> str:
		if self.done is None:
			return ""
		if self.done:
			# Translators: Shown in the comment list for a comment marked as resolved.
			return _("Resolved")
		# Translators: Shown in the comment list for a comment that is still open.
		return _("Open")

	@property
	def kindText(self) -> str:
		if self.kind == KIND_NOTE:
			# Translators: The Excel annotation kind that has no replies, shown in the Type column.
			return _("Note")
		if self.kind == KIND_THREADED:
			# Translators: The Excel annotation kind that supports replies, shown in the Type column.
			return _("Comment thread")
		return ""

	@property
	def supportsReplies(self) -> bool:
		"""Excel notes are a single block of text with no thread behind them."""
		return self.kind != KIND_NOTE

	@property
	def supportsResolve(self) -> bool:
		return self.done is not None and not self.isReply

	@property
	def location(self) -> str:
		"""Where the item sits, as one readable string."""
		if self.cellAddress:
			return f"{self.sheetName}!{self.cellAddress}" if self.sheetName else self.cellAddress
		return str(self.page) if self.page else ""

	@property
	def sortKey(self):
		"""Document order, whichever application the item came from."""
		if self.cellAddress:
			return (self.sheetIndex, self.row, self.column, self.threadPosition)
		return (self.startPos if self.startPos is not None else 0, self.index)


@dataclass
class RevisionItem:
	"""A single tracked change. Word only."""

	index: int
	type: int = 0
	author: str = ""
	date: Optional[datetime.datetime] = None
	text: str = ""
	formatDescription: str = ""
	page: Optional[int] = None
	comObject: Any = field(default=None, repr=False)
	scopeObj: Any = field(default=None, repr=False)

	@property
	def typeText(self) -> str:
		# Translators: Reported for a tracked change whose type Word did not name.
		return REVISION_TYPE_LABELS.get(self.type, _("Unknown change"))
