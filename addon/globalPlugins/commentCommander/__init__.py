# -*- coding: UTF-8 -*-
# Comment Commander: work with Microsoft Word comments efficiently from NVDA.
# Copyright (C) 2026 Sensotec
# This file is covered by the GNU General Public License version 2.

import addonHandler
import api
import globalPluginHandler
import gui
import ui
import wx
from scriptHandler import getLastScriptRepeatCount, script

from . import backend as backendModule
from .common import KIND_NOTE, CommentItem, OfficeAccessError
from .dialogs import CommentsDialog, RevisionsDialog, describeComment
from .settings import (
	CommentCommanderSettingsPanel,
	getConf,
	initConfig,
	registerPanel,
	unregisterPanel,
)

addonHandler.initTranslation()

initConfig()

# Translators: The name of the add-on, used as the input gestures category.
SCRIPT_CATEGORY = _("Comment Commander")

#: How long to wait after closing the dialog before moving the Word caret, so
#: that focus has settled back on the document window first.
JUMP_DELAY_MS = 120

#: Scripts that are meaningless outside a supported Office application.
#: Anything not listed here (opening the settings) stays available everywhere.
OFFICE_ONLY_SCRIPTS = frozenset((
	"script_showComments",
	"script_showRevisions",
	"script_reportComment",
	"script_nextComment",
	"script_previousComment",
	"script_addComment",
	"script_reportCounts",
))


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = SCRIPT_CATEGORY

	def __init__(self):
		super().__init__()
		self._dialogOpen = False
		registerPanel()

	def terminate(self):
		unregisterPanel()
		super().terminate()

	def getScript(self, gesture):
		"""Only claim the commands while a supported application has the focus.

		A global plugin is asked about every gesture no matter which application
		is in front, so without this the commands would answer from anywhere.
		The check is intentionally done here rather than inside each script: a
		script that is never handed out cannot act on the wrong document.
		"""
		script = super().getScript(gesture)
		if script is None:
			return None
		name = getattr(script, "__name__", "")
		if name not in OFFICE_ONLY_SCRIPTS:
			return script
		if not backendModule.isOfficeFocused():
			return self.script_notInOffice
		if name == "script_showRevisions" and not backendModule.isWordFocused():
			# Tracked changes exist in Word only.
			return self.script_revisionsNeedWord
		return script

	def script_notInOffice(self, gesture):
		# No docstring on purpose: that keeps it out of the Input Gestures list.
		ui.message(
			# Translators: Reported when the user is not in a supported Office document.
			_("Not in a Microsoft Word or Excel document")
		)

	def script_revisionsNeedWord(self, gesture):
		# No docstring on purpose: that keeps it out of the Input Gestures list.
		ui.message(
			# Translators: Reported when tracked changes are requested outside Word.
			_("Tracked changes are only available in Microsoft Word")
		)

	# -- Helpers --------------------------------------------------------------

	def _getBackend(self):
		"""Return (backend, document), announcing the problem if there is none."""
		backend, document = backendModule.getBackend()
		if document is None:
			ui.message(
				# Translators: Reported when the user is not in a supported Office document.
				_("Not in a Microsoft Word or Excel document")
			)
			return None, None
		return backend, document

	def _loadComments(self, backend, doc, withDetails=True, withPositions=False):
		try:
			return backend.getComments(
				doc, withDetails=withDetails, withPositions=withPositions
			)
		except OfficeAccessError as e:
			ui.message(str(e))
			return None

	def _announceArrival(self, item):
		"""Speak whatever the user asked to hear after the caret has moved."""
		mode = getConf()["announceAfterJump"]
		if mode == "none":
			return
		if not isinstance(item, CommentItem):
			return
		if mode == "anchor":
			if item.anchorText:
				ui.message(item.anchorText)
			return
		if mode == "comment":
			ui.message(describeComment(item, includeAnchor=False))
			return
		ui.message(describeComment(item, includeAnchor=True))

	def _jumpTo(self, backend, item):
		"""Move the caret to an item once focus is back in the document."""

		def doJump():
			if not backend.goTo(item, selectAnnotatedText=getConf()["selectAnnotatedText"]):
				ui.message(
					# Translators: Reported when the cursor could not be moved to a comment.
					_("Could not move to this item")
				)
				return
			self._announceArrival(item)

		wx.CallLater(JUMP_DELAY_MS, doJump)

	def _showDialog(self, dialogClass, loader):
		"""Load the items, show the dialog, then act on what the user chose."""
		if self._dialogOpen:
			return
		backend, doc = self._getBackend()
		if doc is None:
			return
		items = loader(backend, doc)
		if items is None:
			return
		if not items:
			ui.message(self._emptyMessage(dialogClass))
			return

		self._dialogOpen = True
		chosen = None
		gui.mainFrame.prePopup()
		try:
			dialog = dialogClass(gui.mainFrame, backend, doc, items)
			try:
				if dialog.ShowModal() == wx.ID_OK:
					chosen = dialog.chosenItem
				dialog.saveState()
			finally:
				dialog.Destroy()
		finally:
			gui.mainFrame.postPopup()
			self._dialogOpen = False
		# postPopup has handed focus back to the document, so the caret can move.
		if chosen is not None:
			self._jumpTo(backend, chosen)

	def _emptyMessage(self, dialogClass):
		if dialogClass is RevisionsDialog:
			# Translators: Reported when the document contains no tracked changes.
			return _("No tracked changes in this document")
		# Translators: Reported when the document contains no comments.
		return _("No comments in this document")

	# -- Scripts --------------------------------------------------------------

	@script(
		# Translators: The description of a command, shown in NVDA's Input Gestures dialog.
		description=_("Shows a list of all comments in the current Word document or Excel workbook"),
		gesture="kb:NVDA+shift+;",
	)
	def script_showComments(self, gesture):
		wx.CallAfter(self._showDialog, CommentsDialog, self._loadComments)

	@script(
		# Translators: The description of a command, shown in NVDA's Input Gestures dialog.
		description=_("Shows a list of all tracked changes in the current Microsoft Word document"),
		gesture="kb:NVDA+control+shift+;",
	)
	def script_showRevisions(self, gesture):
		def loader(backend, doc):
			try:
				return backend.getRevisions(doc)
			except OfficeAccessError as e:
				ui.message(str(e))
				return None

		wx.CallAfter(self._showDialog, RevisionsDialog, loader)

	@script(
		# Translators: The description of a command, shown in NVDA's Input Gestures dialog.
		description=_(
			"Reports the comment at the cursor. Press twice to copy it to the clipboard"
		),
		gesture="kb:NVDA+;",
	)
	def script_reportComment(self, gesture):
		backend, doc = self._getBackend()
		if doc is None:
			return
		# Only the anchor positions are needed to find the comment at the caret;
		# the text of every other comment in the document is never used.
		comments = self._loadComments(backend, doc, withDetails=False, withPositions=True)
		if comments is None:
			return
		matches = backend.getCommentsAtSelection(doc, comments)
		if not matches:
			ui.message(
				# Translators: Reported when there is no comment at the cursor position.
				_("No comment here")
			)
			return
		# Include the replies belonging to each matching thread, and read the
		# full text for just those.
		full = []
		for thread in matches:
			full.append(thread)
			full.extend(c for c in comments if c.parentIndex == thread.index)
		text = "\n".join(describeComment(backend.loadDetails(c)) for c in full)
		if getLastScriptRepeatCount() >= 1:
			if api.copyToClip(text):
				# Translators: Reported when the comment was copied to the clipboard.
				ui.message(_("Comment copied to clipboard"))
			else:
				# Translators: Reported when copying to the clipboard failed.
				ui.message(_("Could not copy to clipboard"))
			return
		ui.message(text)

	def _moveToAdjacent(self, forward: bool):
		backend, doc = self._getBackend()
		if doc is None:
			return
		comments = self._loadComments(backend, doc, withDetails=False, withPositions=True)
		if comments is None:
			return
		if not comments:
			# Translators: Reported when the document contains no comments.
			ui.message(_("No comments in this document"))
			return
		item = backend.findAdjacentComment(doc, comments, forward=forward)
		if item is not None:
			# Only the comment we actually landed on needs its text read.
			backend.loadDetails(item)
		if item is None:
			if forward:
				# Translators: Reported when there is no comment after the cursor.
				ui.message(_("No next comment"))
			else:
				# Translators: Reported when there is no comment before the cursor.
				ui.message(_("No previous comment"))
			return
		if not backend.goTo(item, selectAnnotatedText=getConf()["selectAnnotatedText"]):
			# Translators: Reported when the cursor could not be moved to a comment.
			ui.message(_("Could not move to this item"))
			return
		self._announceArrival(item)

	@script(
		# Translators: The description of a command, shown in NVDA's Input Gestures dialog.
		description=_("Moves to the next comment"),
		gesture="kb:NVDA+alt+;",
	)
	def script_nextComment(self, gesture):
		self._moveToAdjacent(forward=True)

	@script(
		# Translators: The description of a command, shown in NVDA's Input Gestures dialog.
		description=_("Moves to the previous comment"),
		gesture="kb:NVDA+alt+shift+;",
	)
	def script_previousComment(self, gesture):
		self._moveToAdjacent(forward=False)

	@script(
		# Translators: The description of a command, shown in NVDA's Input Gestures dialog.
		description=_("Adds a comment on the current selection or cell"),
		gesture="kb:NVDA+control+;",
	)
	def script_addComment(self, gesture):
		backend, doc = self._getBackend()
		if doc is None:
			return

		def ask():
			# Translators: The message in the dialog for typing a new comment.
			message = _("Comment text:")
			with wx.TextEntryDialog(
				gui.mainFrame,
				message,
				# Translators: The title of the dialog for typing a new comment.
				_("New comment"),
				style=wx.TE_MULTILINE | wx.OK | wx.CANCEL,
			) as dlg:
				gui.mainFrame.prePopup()
				try:
					if dlg.ShowModal() != wx.ID_OK:
						return
					text = dlg.GetValue().strip()
				finally:
					gui.mainFrame.postPopup()
			if not text:
				return
			try:
				backend.addComment(doc, text)
			except OfficeAccessError as e:
				ui.message(str(e))
				return
			# Translators: Reported after a new comment was added.
			ui.message(_("Comment added"))

		wx.CallAfter(ask)

	@script(
		# Translators: The description of a command, shown in NVDA's Input Gestures dialog.
		description=_("Reports how many comments the document contains"),
	)
	def script_reportCounts(self, gesture):
		backend, doc = self._getBackend()
		if doc is None:
			return
		comments = self._loadComments(backend, doc) or []
		threads = [c for c in comments if not c.isReply]
		if backendModule.kindOf(backend) == backendModule.EXCEL:
			notes = len([c for c in threads if c.kind == KIND_NOTE])
			ui.message(
				# Translators: Reports a summary of the annotations in an Excel workbook.
				_("{notes} notes, {threads} comment threads").format(
					notes=notes, threads=len(threads) - notes
				)
			)
			return
		resolved = len([c for c in threads if c.done])
		try:
			revisions = backend.getRevisions(doc, countOnly=True)
		except OfficeAccessError:
			revisions = []
		ui.message(
			# Translators: Reports a summary of the comments and tracked changes in the document.
			_("{comments} comments, {resolved} resolved, {revisions} tracked changes").format(
				comments=len(threads), resolved=resolved, revisions=len(revisions)
			)
		)

	@script(
		# Translators: The description of a command, shown in NVDA's Input Gestures dialog.
		description=_("Opens the Comment Commander settings"),
	)
	def script_openSettings(self, gesture):
		wx.CallAfter(
			gui.mainFrame.popupSettingsDialog,
			gui.settingsDialogs.NVDASettingsDialog,
			CommentCommanderSettingsPanel,
		)
