# -*- coding: UTF-8 -*-
# Comment Commander: Microsoft Word object model access layer.
# Copyright (C) 2026 Sensotec
# This file is covered by the GNU General Public License version 2.

"""Everything that talks to Microsoft Word lives here.

The rest of the add-on only ever sees the plain data classes defined below, so
the UI never has to care whether NVDA is driving Word through UIA or through
the legacy object model, nor whether the document uses classic or modern
comments.
"""

from typing import List, Optional

import addonHandler
import api
import winUser
from comtypes import COMError
from logHandler import log

from .common import (
	Capabilities,
	CommentItem,
	OfficeAccessError,
	RevisionItem,
	WordAccessError,
	dispatchFromWindow,
	getFocusAppName,
	hasWindowClassInAncestry,
	normalize,
	safeGet,
	shareDispatchIds,
)

addonHandler.initTranslation()


#: Window classes that expose the Word object model via OBJID_NATIVEOM.
WORD_DOCUMENT_WINDOW_CLASSES = ("_WwG", "_WwN", "_WwO", "_WwB")

#: NVDA application names that are Microsoft Word itself.
WORD_APP_NAMES = frozenset(("winword",))

#: Applications that embed Word as their editor. These only count when a real
#: Word document window is present, so the Outlook message list does not match.
WORD_HOST_APP_NAMES = frozenset(("outlook",))

# Word enumeration values we rely on. Defined locally so we never depend on a
# generated type library being present.
wdCollapseStart = 1
wdActiveEndPageNumber = 3

def isWordFocused() -> bool:
	"""True when Microsoft Word currently has the focus.

	Deliberately cheap and COM-free: this runs on every keypress that matches
	one of the add-on's gestures, so it may not talk to Word.
	"""
	appName = getFocusAppName()
	if appName in WORD_APP_NAMES:
		return True
	if appName in WORD_HOST_APP_NAMES:
		try:
			focus = api.getFocusObject()
		except Exception:
			return False
		# Word hosted inside another application, e.g. editing an Outlook message.
		return hasWindowClassInAncestry(focus, WORD_DOCUMENT_WINDOW_CLASSES)
	return False


def _documentFromWindowObject(windowObject):
	"""Given a Word Window object, return its Document."""
	for attempt in ("Document", None):
		try:
			if attempt:
				return windowObject.Document
			return windowObject.Application.ActiveDocument
		except (COMError, AttributeError, NameError):
			continue
	return None


def _documentFromNVDAObjects():
	"""Ask NVDA's own Word objects for the document, walking up from focus."""
	obj = api.getFocusObject()
	depth = 0
	while obj is not None and depth < 8:
		try:
			doc = getattr(obj, "WinwordDocumentObject", None)
		except Exception:
			# Fetching this property can raise anything at all if Word is busy.
			doc = None
		if doc is not None:
			return doc
		obj = obj.parent
		depth += 1
	return None


def _documentFromWindowHandles():
	"""Locate a Word document window relative to the focused window."""
	focus = api.getFocusObject()
	hwnd = getattr(focus, "windowHandle", 0)
	if not hwnd:
		return None
	# Walk up from the focused window; also remember the top level window so we
	# can search downwards when focus is on the ribbon or a task pane.
	topLevel = hwnd
	candidate = hwnd
	depth = 0
	while candidate and depth < 12:
		try:
			className = winUser.getClassName(candidate)
		except Exception:
			className = ""
		if className in WORD_DOCUMENT_WINDOW_CLASSES:
			windowObject = dispatchFromWindow(candidate)
			if windowObject is not None:
				doc = _documentFromWindowObject(windowObject)
				if doc is not None:
					return doc
		topLevel = candidate
		candidate = winUser.getAncestor(candidate, winUser.GA_PARENT)
		depth += 1
	# Nothing in the ancestry: search the frame for a document window.
	try:
		import windowUtils

		for className in WORD_DOCUMENT_WINDOW_CLASSES:
			try:
				childHwnd = windowUtils.findDescendantWindow(topLevel, className=className)
			except LookupError:
				continue
			windowObject = dispatchFromWindow(childHwnd)
			if windowObject is not None:
				doc = _documentFromWindowObject(windowObject)
				if doc is not None:
					return doc
	except Exception:
		log.debugWarning("Comment Commander: descendant window search failed", exc_info=True)
	return None


def getDocumentObject(requireFocus: bool = True):
	"""Return the Word Document object for the current context, or None.

	Tries NVDA's own object model handle first (cheapest and most accurate),
	then the window handles around focus.

	Both routes start from the focused window, so they can never reach a Word
	document in the background. That matters: an earlier version also fell back
	to the running object table, which happily returned Word's active document
	even when the user was in a completely different application.
	"""
	if requireFocus and not isWordFocused():
		return None
	for getter in (
		_documentFromNVDAObjects,
		_documentFromWindowHandles,
	):
		try:
			doc = getter()
		except Exception:
			log.debugWarning(f"Comment Commander: {getter.__name__} failed", exc_info=True)
			continue
		if doc is not None:
			return doc
	return None


def isWordAvailable() -> bool:
	return getDocumentObject() is not None


def getDocumentName(doc) -> str:
	try:
		return doc.Name
	except COMError:
		return ""


def getPage(item) -> Optional[int]:
	"""Fetch and cache the page number for one item.

	Kept out of the bulk read on purpose: Information() forces Word to
	paginate, so asking for every comment at once is what made large documents
	slow. Asking only for the row the user is actually on costs nothing
	noticeable.
	"""
	if item.page is not None:
		return item.page
	rangeObj = item.scopeObj
	if rangeObj is None:
		rangeObj = _targetRange(item)
		item.scopeObj = rangeObj
	if rangeObj is None:
		return None
	item.page = safeGet(lambda: int(rangeObj.Information(wdActiveEndPageNumber)))
	return item.page


def loadDetails(item: CommentItem) -> CommentItem:
	"""Fill in the descriptive fields of a comment read in positions-only mode."""
	if item.detailsLoaded:
		return item
	com = item.comObject
	if com is None:
		return item
	capabilities = Capabilities()
	scope = item.scopeObj
	if scope is None:
		scope = safeGet(lambda: com.Scope)
		item.scopeObj = scope
	item.author = safeGet(lambda: com.Author, "") or ""
	item.date = safeGet(lambda: com.Date)
	item.text = normalize(safeGet(lambda: com.Range.Text, ""))
	item.anchorText = normalize(safeGet(lambda: scope.Text, "") if scope is not None else "")
	item.done = capabilities.read("done", lambda: bool(com.Done))
	item.detailsLoaded = True
	return item


def getComments(doc, withDetails: bool = True, withPositions: bool = False) -> List[CommentItem]:
	"""Read every comment in the document, replies included.

	Replies are returned directly after the comment they belong to, so the list
	reads as a set of threads even though it is flat.

	@param withDetails: read author, date and text. The list dialog needs these;
		the navigation commands do not.
	@param withPositions: read and cache the character offsets of each anchor.
		Needed to work out which comment the caret is in.
	"""
	try:
		collection = doc.Comments
		count = int(collection.Count)
	except COMError as e:
		raise WordAccessError(
			# Translators: Reported when the comments could not be read from Word.
			_("The comments in this document could not be read.")
		) from e
	if not count:
		return []

	capabilities = Capabilities()
	commentTemplate = None
	rangeTemplate = None
	raw = {}
	order = []
	for i in range(1, count + 1):
		try:
			com = collection.Item(i)
		except COMError:
			log.debugWarning(f"Comment Commander: could not read comment {i}", exc_info=True)
			continue
		shareDispatchIds(commentTemplate, com)
		if commentTemplate is None:
			commentTemplate = com

		scope = safeGet(lambda: com.Scope)
		if scope is not None:
			shareDispatchIds(rangeTemplate, scope)
			if rangeTemplate is None:
				rangeTemplate = scope

		parentIndex = None
		ancestor = capabilities.read("ancestor", lambda: com.Ancestor)
		if ancestor is not None:
			shareDispatchIds(commentTemplate, ancestor)
			parentIndex = safeGet(lambda: int(ancestor.Index))

		item = CommentItem(
			index=i,
			parentIndex=parentIndex,
			comObject=com,
			scopeObj=scope,
		)
		if withPositions and scope is not None:
			item.startPos = safeGet(lambda: int(scope.Start))
			item.endPos = safeGet(lambda: int(scope.End))
		if withDetails:
			item.author = safeGet(lambda: com.Author, "") or ""
			item.date = safeGet(lambda: com.Date)
			# Comment.Range is a Word Range just like Scope, so it can reuse the
			# same dispatch-ID template.
			body = safeGet(lambda: com.Range)
			if body is not None:
				shareDispatchIds(rangeTemplate, body)
				if rangeTemplate is None:
					rangeTemplate = body
			item.text = normalize(safeGet(lambda: body.Text, "") if body is not None else "")
			item.anchorText = normalize(safeGet(lambda: scope.Text, "") if scope is not None else "")
			item.done = capabilities.read("done", lambda: bool(com.Done))
			item.detailsLoaded = True
		raw[i] = item
		order.append(i)

	# Reply counts come from the parent links we already have, rather than from
	# a Replies.Count call per comment.
	for i in order:
		parent = raw[i].parentIndex
		if parent is not None and parent in raw:
			raw[parent].replyCount += 1

	# Group replies under their parent while preserving document order.
	items: List[CommentItem] = []
	consumed = set()
	for i in order:
		item = raw[i]
		if item.isReply or i in consumed:
			continue
		item.threadPosition = 0
		items.append(item)
		consumed.add(i)
		position = 1
		for j in order:
			child = raw[j]
			if child.parentIndex == i and j not in consumed:
				child.threadPosition = position
				items.append(child)
				consumed.add(j)
				position += 1
	# Anything whose parent we failed to resolve still deserves to be listed.
	for i in order:
		if i not in consumed:
			items.append(raw[i])
	return items


def getRevisions(doc, countOnly: bool = False) -> List[RevisionItem]:
	"""Read every tracked change in the document.

	Page numbers are left out here too and fetched per row by L{getPage}.
	"""
	try:
		collection = doc.Revisions
		count = int(collection.Count)
	except COMError as e:
		raise WordAccessError(
			# Translators: Reported when the tracked changes could not be read from Word.
			_("The tracked changes in this document could not be read.")
		) from e
	if not count:
		return []
	if countOnly:
		return [RevisionItem(index=i) for i in range(1, count + 1)]

	capabilities = Capabilities()
	revisionTemplate = None
	rangeTemplate = None
	items: List[RevisionItem] = []
	for i in range(1, count + 1):
		try:
			com = collection.Item(i)
		except COMError:
			log.debugWarning(f"Comment Commander: could not read revision {i}", exc_info=True)
			continue
		shareDispatchIds(revisionTemplate, com)
		if revisionTemplate is None:
			revisionTemplate = com
		rangeObj = safeGet(lambda: com.Range)
		if rangeObj is not None:
			shareDispatchIds(rangeTemplate, rangeObj)
			if rangeTemplate is None:
				rangeTemplate = rangeObj
		items.append(
			RevisionItem(
				index=i,
				type=safeGet(lambda: int(com.Type), 0) or 0,
				author=safeGet(lambda: com.Author, "") or "",
				date=safeGet(lambda: com.Date),
				text=normalize(safeGet(lambda: rangeObj.Text, "") if rangeObj is not None else ""),
				formatDescription=capabilities.read(
					"done", lambda: normalize(com.FormatDescription or "")
				) or "",
				comObject=com,
				scopeObj=rangeObj,
			)
		)
	return items


def _targetRange(item):
	"""The document range an item points at."""
	if item.scopeObj is not None:
		return item.scopeObj
	if isinstance(item, CommentItem):
		# A reply has no scope of its own; use the thread's anchor.
		item.scopeObj = safeGet(lambda: item.comObject.Scope)
	else:
		item.scopeObj = safeGet(lambda: item.comObject.Range)
	return item.scopeObj


def goTo(item, selectAnnotatedText: bool = False) -> bool:
	"""Move the Word caret to the item. Returns True on success."""
	rangeObj = _targetRange(item)
	if rangeObj is None:
		return False
	try:
		if selectAnnotatedText:
			rangeObj.Select()
		else:
			target = rangeObj.Duplicate
			target.Collapse(wdCollapseStart)
			target.Select()
	except COMError:
		log.debugWarning("Comment Commander: could not select range", exc_info=True)
		return False
	# Bring the document window forward in case focus drifted to a task pane.
	try:
		item.comObject.Application.ActiveWindow.Activate()
	except COMError:
		pass
	return True


def getSelectionBounds(doc):
	"""Return (start, end) of the current selection, or None."""
	try:
		selection = doc.ActiveWindow.Selection
		return int(selection.Start), int(selection.End)
	except COMError:
		try:
			selection = doc.Application.Selection
			return int(selection.Start), int(selection.End)
		except COMError:
			return None


def _bounds(item):
	"""Character offsets of an item's anchor, reading them only if not cached."""
	if item.startPos is None or item.endPos is None:
		scope = _targetRange(item)
		if scope is None:
			return None, None
		item.startPos = safeGet(lambda: int(scope.Start))
		item.endPos = safeGet(lambda: int(scope.End))
	return item.startPos, item.endPos


def getCommentsAtSelection(doc, comments: List[CommentItem]) -> List[CommentItem]:
	"""Return the comment threads whose anchor overlaps the caret."""
	bounds = getSelectionBounds(doc)
	if bounds is None:
		return []
	selStart, selEnd = bounds
	matches = []
	for item in comments:
		if item.isReply:
			continue
		start, end = _bounds(item)
		if start is None or end is None:
			continue
		# Treat a zero length anchor as covering the character it sits on.
		if start <= selEnd and end >= selStart:
			matches.append(item)
	return matches


def findAdjacentComment(doc, comments: List[CommentItem], forward: bool = True) -> Optional[CommentItem]:
	"""Find the next or previous comment thread relative to the caret."""
	bounds = getSelectionBounds(doc)
	if bounds is None:
		return None
	selStart, selEnd = bounds
	best = None
	bestStart = None
	for item in comments:
		if item.isReply:
			continue
		start, end = _bounds(item)
		if start is None or end is None:
			continue
		if start <= selStart and end >= selEnd:
			# The caret is already inside this comment, so it is neither the
			# next nor the previous one. Without this, moving backwards out of
			# a comment would keep landing on the comment you are standing in.
			continue
		if forward and start > selEnd:
			if bestStart is None or start < bestStart:
				best, bestStart = item, start
		elif not forward and start < selStart:
			if bestStart is None or start > bestStart:
				best, bestStart = item, start
	return best


def addReply(item: CommentItem, text: str) -> None:
	"""Add a reply to a comment thread.

	comtypes' dynamic dispatch cannot pass named arguments, so the range has to
	be supplied positionally even though Word treats it as optional here.
	"""
	target = item.comObject
	if item.isReply:
		# Replies live on the thread, so walk up to the root comment first.
		ancestor = safeGet(lambda: target.Ancestor)
		if ancestor is not None:
			target = ancestor
	lastError = None
	for rangeGetter in (lambda: target.Scope, lambda: target.Range):
		try:
			target.Replies.Add(rangeGetter(), text)
			return
		except COMError as e:
			lastError = e
	raise WordAccessError(
		# Translators: Reported when Word refused to add a reply to a comment.
		_("Word could not add this reply. The document may be protected, or replies may not be supported here.")
	) from lastError


def setCommentText(item: CommentItem, text: str) -> None:
	"""Replace the body of a comment."""
	try:
		item.comObject.Range.Text = text
	except COMError as e:
		raise WordAccessError(
			# Translators: Reported when Word refused to change the text of a comment.
			_("Word could not change this comment. The document may be protected.")
		) from e


def setCommentDone(item: CommentItem, done: bool) -> None:
	"""Mark a comment thread as resolved or reopen it."""
	try:
		item.comObject.Done = done
	except COMError as e:
		raise WordAccessError(
			# Translators: Reported when Word refused to resolve or reopen a comment.
			_("Word could not change the status of this comment. Your version of Word may not support resolving comments.")
		) from e


def deleteItem(item) -> None:
	"""Delete a comment, or reject nothing - revisions use accept/reject."""
	try:
		item.comObject.Delete()
	except COMError as e:
		raise WordAccessError(
			# Translators: Reported when Word refused to delete a comment.
			_("Word could not delete this item. The document may be protected.")
		) from e


def acceptRevision(item: RevisionItem) -> None:
	try:
		item.comObject.Accept()
	except COMError as e:
		raise WordAccessError(
			# Translators: Reported when Word refused to accept a tracked change.
			_("Word could not accept this change. The document may be protected.")
		) from e


def rejectRevision(item: RevisionItem) -> None:
	try:
		item.comObject.Reject()
	except COMError as e:
		raise WordAccessError(
			# Translators: Reported when Word refused to reject a tracked change.
			_("Word could not reject this change. The document may be protected.")
		) from e


def addComment(doc, text: str) -> None:
	"""Add a new comment on the current selection."""
	bounds = getSelectionBounds(doc)
	if bounds is None:
		raise WordAccessError(
			# Translators: Reported when the caret position in Word could not be determined.
			_("The cursor position in Word could not be determined.")
		)
	try:
		selection = doc.ActiveWindow.Selection
		doc.Comments.Add(selection.Range, text)
	except COMError as e:
		raise WordAccessError(
			# Translators: Reported when Word refused to add a new comment.
			_("Word could not add a comment here. The document may be protected.")
		) from e
