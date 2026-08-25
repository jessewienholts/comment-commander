# -*- coding: UTF-8 -*-
# Comment Commander: Microsoft Excel object model access layer.
# Copyright (C) 2026 Sensotec
# This file is covered by the GNU General Public License version 2.

"""Excel backend.

Excel has two separate kinds of cell annotation and exposes them through two
separate collections:

* **Notes** are the classic yellow sticky notes, reached through
  ``Worksheet.Comments``. They have an author and text, but no date, no replies
  and no resolved state.
* **Comment threads** are the modern ones, reached through
  ``Worksheet.CommentsThreaded``. They add a date and replies. Their author is
  an object rather than a string.

Both end up in the same list. Microsoft's own documentation is contradictory
about whether ``CommentsThreaded`` also returns the legacy notes, so the two
collections are merged and then deduplicated by cell address: correct either
way, rather than betting on one reading.
"""

from typing import List, Optional

import addonHandler
import api
from comtypes import COMError
from logHandler import log

from .common import (
	KIND_NOTE,
	KIND_THREADED,
	Capabilities,
	CommentItem,
	OfficeAccessError,
	flagAsMethod,
	dispatchFromWindow,
	getFocusAppName,
	hasWindowClassInAncestry,
	normalize,
	readValue,
	safeGet,
	shareDispatchIds,
)

addonHandler.initTranslation()

#: The grid window exposes the Excel object model via OBJID_NATIVEOM.
EXCEL_WINDOW_CLASSES = ("EXCEL7",)

#: NVDA application names that are Microsoft Excel.
EXCEL_APP_NAMES = frozenset(("excel",))

#: Excel enumeration values we rely on.
xlA1 = 1

#: Refuse to enumerate absurd workbooks rather than freezing NVDA.
MAX_SHEETS = 200


def isExcelFocused() -> bool:
	"""True when Microsoft Excel currently has the focus.

	Deliberately cheap and COM-free: this runs on every keypress that matches
	one of the add-on's gestures, so it may not talk to Excel.
	"""
	return getFocusAppName() in EXCEL_APP_NAMES


def _workbookFromWindowObject(windowObject):
	"""Given an Excel Window object, return the Workbook it shows."""
	for getter in (
		lambda: windowObject.Application.ActiveWorkbook,
		lambda: windowObject.Parent,
		lambda: windowObject.Workbook,
	):
		book = safeGet(getter)
		# Probe for something only a Workbook has, so we never hand back the
		# Application object by mistake.
		if book is not None and safeGet(lambda: book.Worksheets) is not None:
			return book
	return None


def _workbookFromNVDAObjects():
	"""Ask NVDA's own Excel objects for the workbook, walking up from focus."""
	try:
		obj = api.getFocusObject()
	except Exception:
		return None
	depth = 0
	while obj is not None and depth < 8:
		for attribute in ("excelWorksheetObject", "excelWindowObject"):
			try:
				handle = getattr(obj, attribute, None)
			except Exception:
				handle = None
			if handle is None:
				continue
			if attribute == "excelWorksheetObject":
				book = safeGet(lambda: handle.Parent)
				if book is not None and safeGet(lambda: book.Worksheets) is not None:
					return book
			else:
				book = _workbookFromWindowObject(handle)
				if book is not None:
					return book
		obj = obj.parent
		depth += 1
	return None


def _workbookFromWindowHandles():
	"""Locate the Excel grid window relative to the focused window."""
	try:
		focus = api.getFocusObject()
	except Exception:
		return None
	hwnd = getattr(focus, "windowHandle", 0)
	if not hwnd:
		return None
	import winUser

	topLevel = hwnd
	candidate = hwnd
	depth = 0
	while candidate and depth < 12:
		try:
			className = winUser.getClassName(candidate)
		except Exception:
			className = ""
		if className in EXCEL_WINDOW_CLASSES:
			windowObject = dispatchFromWindow(candidate)
			if windowObject is not None:
				book = _workbookFromWindowObject(windowObject)
				if book is not None:
					return book
		topLevel = candidate
		candidate = winUser.getAncestor(candidate, winUser.GA_PARENT)
		depth += 1
	# Focus may be on the ribbon or a task pane, so search the frame downwards.
	try:
		import windowUtils

		for className in EXCEL_WINDOW_CLASSES:
			try:
				childHwnd = windowUtils.findDescendantWindow(topLevel, className=className)
			except LookupError:
				continue
			windowObject = dispatchFromWindow(childHwnd)
			if windowObject is not None:
				book = _workbookFromWindowObject(windowObject)
				if book is not None:
					return book
	except Exception:
		log.debugWarning("Comment Commander: Excel descendant search failed", exc_info=True)
	return None


def getDocumentObject(requireFocus: bool = True):
	"""Return the Excel Workbook for the current context, or None.

	Like the Word backend, both routes start from the focused window, so this
	can never reach a workbook sitting in the background.
	"""
	if requireFocus and not isExcelFocused():
		return None
	for getter in (_workbookFromNVDAObjects, _workbookFromWindowHandles):
		try:
			book = getter()
		except Exception:
			log.debugWarning(f"Comment Commander: {getter.__name__} failed", exc_info=True)
			continue
		if book is not None:
			return book
	return None


def isAvailable() -> bool:
	return getDocumentObject() is not None


def getDocumentName(book) -> str:
	return safeGet(lambda: book.Name, "") or ""


def _cellOf(com):
	"""The cell an annotation belongs to.

	Microsoft documents Parent only as "the parent object", so rather than trust
	it blindly we check that whatever comes back actually looks like a range.
	Notes are drawing shapes, which gives a second route when Parent disappoints.
	"""
	for getter in (
		lambda: com.Parent,
		lambda: com.Shape.TopLeftCell,
	):
		cell = safeGet(getter)
		if cell is not None and safeGet(lambda: cell.Address) is not None:
			return cell
	return None


def _addressOf(cell) -> str:
	"""The cell address without dollar signs, e.g. B7."""
	address = readValue(lambda: cell.Address, "")
	if isinstance(address, str):
		return address.replace("$", "")
	# Some Excel builds expose Address as a method needing explicit arguments.
	address = safeGet(lambda: cell.Address(False, False, xlA1), "")
	return address.replace("$", "") if isinstance(address, str) else ""


def _cellText(cell) -> str:
	"""What the commented cell shows, used as the anchor text."""
	return normalize(safeGet(lambda: cell.Text, "") or safeGet(lambda: cell.Value, "") or "")


def _authorName(com, threaded: bool) -> str:
	"""Notes hand back a plain string; threads hand back an Author object."""
	if not threaded:
		return normalize(safeGet(lambda: com.Author, "") or "")
	author = safeGet(lambda: com.Author)
	if author is None:
		return ""
	name = safeGet(lambda: author.Name)
	if isinstance(name, str):
		return normalize(name)
	return normalize(author if isinstance(author, str) else "")


def _readThreaded(sheet, sheetIndex, sheetName, counter, byAddress, capabilities):
	"""Read the modern comment threads on one worksheet, replies included."""
	collection = safeGet(lambda: sheet.CommentsThreaded)
	if collection is None:
		return
	count = safeGet(lambda: int(collection.Count), 0) or 0
	commentTemplate = None
	cellTemplate = None
	for i in range(1, count + 1):
		com = safeGet(lambda: collection.Item(i))
		if com is None:
			continue
		shareDispatchIds(commentTemplate, com)
		if commentTemplate is None:
			commentTemplate = com
		cell = _cellOf(com)
		if cell is None:
			continue
		shareDispatchIds(cellTemplate, cell)
		if cellTemplate is None:
			cellTemplate = cell
		address = _addressOf(cell)
		counter["n"] += 1
		item = CommentItem(
			index=counter["n"],
			kind=KIND_THREADED,
			author=_authorName(com, True),
			date=capabilities.read("date", lambda: com.Date),
			text=normalize(readValue(lambda: com.Text, "")),
			anchorText=_cellText(cell),
			sheetName=sheetName,
			sheetIndex=sheetIndex,
			cellAddress=address,
			row=safeGet(lambda: int(cell.Row), 0) or 0,
			column=safeGet(lambda: int(cell.Column), 0) or 0,
			comObject=com,
			scopeObj=cell,
			detailsLoaded=True,
		)
		byAddress[(sheetIndex, address)] = [item]

		replies = capabilities.read("replies", lambda: com.Replies)
		if replies is None:
			continue
		replyCount = safeGet(lambda: int(replies.Count), 0) or 0
		item.replyCount = replyCount
		replyTemplate = None
		for r in range(1, replyCount + 1):
			reply = safeGet(lambda: replies.Item(r))
			if reply is None:
				continue
			shareDispatchIds(replyTemplate or commentTemplate, reply)
			if replyTemplate is None:
				replyTemplate = reply
			counter["n"] += 1
			byAddress[(sheetIndex, address)].append(
				CommentItem(
					index=counter["n"],
					kind=KIND_THREADED,
					parentIndex=item.index,
					threadPosition=r,
					author=_authorName(reply, True),
					date=capabilities.read("date", lambda: reply.Date),
					text=normalize(readValue(lambda: reply.Text, "")),
					anchorText=item.anchorText,
					sheetName=sheetName,
					sheetIndex=sheetIndex,
					cellAddress=address,
					row=item.row,
					column=item.column,
					comObject=reply,
					scopeObj=cell,
					detailsLoaded=True,
				)
			)


def _readNotes(sheet, sheetIndex, sheetName, counter, byAddress):
	"""Read the classic notes on one worksheet, skipping cells already covered."""
	collection = safeGet(lambda: sheet.Comments)
	if collection is None:
		return
	count = safeGet(lambda: int(collection.Count), 0) or 0
	noteTemplate = None
	cellTemplate = None
	for i in range(1, count + 1):
		com = safeGet(lambda: collection.Item(i))
		if com is None:
			continue
		shareDispatchIds(noteTemplate, com)
		if noteTemplate is None:
			noteTemplate = com
		cell = _cellOf(com)
		if cell is None:
			continue
		shareDispatchIds(cellTemplate, cell)
		if cellTemplate is None:
			cellTemplate = cell
		address = _addressOf(cell)
		if (sheetIndex, address) in byAddress:
			# Already listed as a thread. Excel's documentation suggests the
			# threaded collection can include legacy notes, so this keeps a
			# single cell from appearing twice.
			continue
		counter["n"] += 1
		byAddress[(sheetIndex, address)] = [
			CommentItem(
				index=counter["n"],
				kind=KIND_NOTE,
				author=_authorName(com, False),
				text=normalize(readValue(lambda: com.Text, "")),
				anchorText=_cellText(cell),
				sheetName=sheetName,
				sheetIndex=sheetIndex,
				cellAddress=address,
				row=safeGet(lambda: int(cell.Row), 0) or 0,
				column=safeGet(lambda: int(cell.Column), 0) or 0,
				comObject=com,
				scopeObj=cell,
				detailsLoaded=True,
			)
		]


def getComments(book, withDetails: bool = True, withPositions: bool = False) -> List[CommentItem]:
	"""Read every note and comment thread in the workbook.

	Excel gives up the cell address as part of the same read, so unlike Word
	there is nothing worth deferring; both flags are accepted only so the two
	backends can be called identically.
	"""
	try:
		sheets = book.Worksheets
		sheetCount = int(sheets.Count)
	except COMError as e:
		raise OfficeAccessError(
			# Translators: Reported when the comments could not be read from Excel.
			_("The comments in this workbook could not be read.")
		) from e
	if sheetCount > MAX_SHEETS:
		log.debugWarning(f"Comment Commander: workbook has {sheetCount} sheets, reading first {MAX_SHEETS}")
		sheetCount = MAX_SHEETS

	capabilities = Capabilities()
	counter = {"n": 0}
	items: List[CommentItem] = []
	for sheetIndex in range(1, sheetCount + 1):
		sheet = safeGet(lambda: sheets.Item(sheetIndex))
		if sheet is None:
			continue
		sheetName = safeGet(lambda: sheet.Name, "") or ""
		byAddress = {}
		_readThreaded(sheet, sheetIndex, sheetName, counter, byAddress, capabilities)
		_readNotes(sheet, sheetIndex, sheetName, counter, byAddress)
		# Cell order within the sheet, thread members kept together.
		for key in sorted(byAddress, key=lambda k: (byAddress[k][0].row, byAddress[k][0].column)):
			items.extend(byAddress[key])
	return items


def getRevisions(book, countOnly: bool = False):
	"""Excel has no per-cell tracked changes comparable to Word's revisions."""
	return []


def loadDetails(item: CommentItem) -> CommentItem:
	"""Everything is read up front for Excel, so there is nothing left to do."""
	return item


def getPage(item) -> Optional[int]:
	"""Excel locates items by sheet and cell rather than by page."""
	return None


def goTo(item, selectAnnotatedText: bool = False) -> bool:
	"""Select the commented cell, activating its sheet first if needed."""
	cell = item.scopeObj
	if cell is None:
		return False
	try:
		sheet = safeGet(lambda: cell.Worksheet)
		if sheet is not None:
			# Range.Select only works on the active sheet.
			sheet.Activate()
		cell.Select()
	except COMError:
		log.debugWarning("Comment Commander: could not select cell", exc_info=True)
		return False
	return True


def _activeCell(book):
	"""Return (sheetIndex, row, column) of the active cell, or None."""
	cell = safeGet(lambda: book.Application.ActiveCell)
	if cell is None:
		return None
	row = safeGet(lambda: int(cell.Row))
	column = safeGet(lambda: int(cell.Column))
	sheetIndex = safeGet(lambda: int(cell.Worksheet.Index), 0) or 0
	if row is None or column is None:
		return None
	return (sheetIndex, row, column)


def getCommentsAtSelection(book, comments: List[CommentItem]) -> List[CommentItem]:
	"""Return the annotation on the active cell, if there is one."""
	position = _activeCell(book)
	if position is None:
		return []
	return [
		item
		for item in comments
		if not item.isReply
		and (item.sheetIndex, item.row, item.column) == position
	]


def findAdjacentComment(book, comments: List[CommentItem], forward: bool = True) -> Optional[CommentItem]:
	"""Find the next or previous annotated cell relative to the active cell."""
	position = _activeCell(book)
	if position is None:
		return None
	threads = [i for i in comments if not i.isReply]
	best = None
	for item in threads:
		key = (item.sheetIndex, item.row, item.column)
		if key == position:
			continue
		if forward and key > position:
			if best is None or key < (best.sheetIndex, best.row, best.column):
				best = item
		elif not forward and key < position:
			if best is None or key > (best.sheetIndex, best.row, best.column):
				best = item
	return best


def addReply(item: CommentItem, text: str) -> None:
	"""Add a reply to a comment thread. Notes cannot hold replies."""
	if item.kind == KIND_NOTE:
		raise OfficeAccessError(
			# Translators: Reported when the user tries to reply to an Excel note.
			_("Notes cannot have replies. Convert it to a comment in Excel first.")
		)
	target = item.comObject
	if item.isReply:
		# Replies belong to the thread, so reply to the cell's thread instead.
		target = safeGet(lambda: item.scopeObj.CommentThreaded) or target
	try:
		flagAsMethod(target, "AddReply")
		target.AddReply(text)
	except (COMError, TypeError) as e:
		raise OfficeAccessError(
			# Translators: Reported when Excel refused to add a reply to a comment.
			_("Excel could not add this reply. The sheet may be protected.")
		) from e


def setCommentText(item: CommentItem, text: str) -> None:
	"""Replace the body of a note or comment."""
	try:
		# Text is a method on both Excel annotation kinds, and it must be flagged
		# as one: an unflagged lookup would property-get it, which wipes the text.
		flagAsMethod(item.comObject, "Text")
		item.comObject.Text(text)
	except (COMError, TypeError) as e:
		raise OfficeAccessError(
			# Translators: Reported when Excel refused to change the text of a comment.
			_("Excel could not change this comment. The sheet may be protected.")
		) from e


def setCommentDone(item: CommentItem, done: bool) -> None:
	"""Excel does not expose the resolved state through the object model."""
	raise OfficeAccessError(
		# Translators: Reported when the user tries to resolve an Excel comment.
		_("Excel does not allow comments to be resolved from outside the application.")
	)


def deleteItem(item) -> None:
	try:
		item.comObject.Delete()
	except COMError as e:
		raise OfficeAccessError(
			# Translators: Reported when Excel refused to delete a comment.
			_("Excel could not delete this item. The sheet may be protected.")
		) from e


def addComment(book, text: str) -> None:
	"""Add a comment thread on the active cell, falling back to a note."""
	cell = safeGet(lambda: book.Application.ActiveCell)
	if cell is None:
		raise OfficeAccessError(
			# Translators: Reported when the active cell in Excel could not be determined.
			_("The active cell in Excel could not be determined.")
		)
	# Both take an optional Text argument, so they must be flagged as methods
	# before use; a bare attribute lookup would run them and leave an empty
	# comment behind. Threaded first, falling back to a classic note.
	flagAsMethod(cell, "AddCommentThreaded", "AddComment")
	lastError = None
	for adder in (
		lambda: cell.AddCommentThreaded(text),
		lambda: cell.AddComment(text),
	):
		try:
			adder()
			return
		except (COMError, TypeError) as e:
			lastError = e
	raise OfficeAccessError(
		# Translators: Reported when Excel refused to add a new comment.
		_("Excel could not add a comment here. The sheet may be protected, or the cell already has one.")
	) from lastError


def acceptRevision(item) -> None:
	raise OfficeAccessError(
		# Translators: Reported when tracked changes are requested in Excel.
		_("Microsoft Excel does not support tracked changes in the way Word does.")
	)


rejectRevision = acceptRevision
