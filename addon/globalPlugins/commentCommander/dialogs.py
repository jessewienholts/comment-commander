# -*- coding: UTF-8 -*-
# Comment Commander: the browse dialogs.
# Copyright (C) 2026 Sensotec
# This file is covered by the GNU General Public License version 2.

"""The list dialogs.

The layout deliberately mirrors the "list of ..." dialogs screen reader users
already know: a filter field, then the list, then everything else. Tab from the
filter lands straight on the list, and Enter on a list item jumps to it in the
document and closes the dialog.
"""

import datetime
from typing import List, Optional

import addonHandler
import gui
import wx
from gui import guiHelper
from logHandler import log

from . import backend as backendModule
from .common import KIND_NOTE, KIND_THREADED, CommentItem, OfficeAccessError
from .settings import getConf

addonHandler.initTranslation()

#: Longest text we put in a list column before truncating.
COLUMN_TEXT_LIMIT = 120

# Translators: Shown in the author filter to include every author.
ALL_AUTHORS = _("All authors")


def formatDate(value: Optional[datetime.datetime]) -> str:
	if not value:
		return ""
	try:
		return value.strftime("%x %H:%M")
	except (ValueError, AttributeError):
		return str(value)


def describeComment(item: CommentItem, includeAnchor: bool = True) -> str:
	"""A spoken description of a comment, used by the reporting scripts.

	The comment itself comes first and the text it is attached to last: what the
	reviewer wrote is the point, where it sits is the follow-up.
	"""
	if item.isReply:
		# Translators: Announced before a reply within a comment thread.
		first = "{label}: {text}".format(label=_("Reply"), text=item.text)
	else:
		first = item.text
	meta = []
	if item.author:
		# Translators: Announced with the author of a comment. {author} is a name.
		meta.append(_("by {author}").format(author=item.author))
	if item.date:
		# Translators: Announced with the date of a comment. {date} is a date and time.
		meta.append(_("on {date}").format(date=formatDate(item.date)))
	if item.done:
		# Translators: Announced for a comment that has been resolved.
		meta.append(_("resolved"))
	if item.replyCount:
		# Translators: Announced with the number of replies in a comment thread.
		meta.append(
			ngettext("{count} reply", "{count} replies", item.replyCount).format(count=item.replyCount)
		)
	if meta:
		first = "{text}, {meta}".format(text=first, meta=", ".join(meta))
	lines = [first]
	if includeAnchor and item.anchorText:
		# Translators: Announced before the document text a comment is attached to.
		lines.append(_("Annotated text: {text}").format(text=item.anchorText))
	return "\n".join(lines)


class BaseItemsDialog(wx.Dialog):
	"""Shared behaviour for the comments and tracked changes lists."""

	#: Set to the item the user activated, or None if they cancelled.
	chosenItem = None

	def __init__(self, parent, backend, doc, items, title):
		conf = getConf()
		self.backend = backend
		#: Excel lists cells and two annotation kinds; Word lists pages and status.
		self.isExcel = backendModule.kindOf(backend) == backendModule.EXCEL
		super().__init__(
			parent,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX,
		)
		self.doc = doc
		self.allItems = items
		self.filteredItems: List = []
		self._suppressDetailUpdate = False

		mainSizer = wx.BoxSizer(wx.VERTICAL)
		helper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)

		# Translators: The label of the filter field in the item list dialogs.
		filterLabel = _("&Filter by text or author:")
		self.filterCtrl = helper.addLabeledControl(
			filterLabel, wx.TextCtrl, style=wx.TE_PROCESS_ENTER
		)
		self.filterCtrl.Bind(wx.EVT_TEXT, self.onFilterChanged)
		self.filterCtrl.Bind(wx.EVT_KEY_DOWN, self.onFilterKeyDown)
		if conf["rememberFilter"] and conf["lastFilterText"]:
			self.filterCtrl.SetValue(conf["lastFilterText"])

		self.itemsList = helper.addLabeledControl(
			self.getListLabel(),
			wx.ListCtrl,
			style=wx.LC_REPORT | wx.LC_SINGLE_SEL,
			size=(conf["dialogWidth"] - 60, 260),
		)
		for index, (label, width) in enumerate(self.getColumns()):
			self.itemsList.InsertColumn(index, label, width=width)
		self.itemsList.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onItemActivated)
		self.itemsList.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onSelectionChanged)
		self.itemsList.Bind(wx.EVT_LIST_ITEM_FOCUSED, self.onSelectionChanged)
		self.itemsList.Bind(wx.EVT_KEY_DOWN, self.onListKeyDown)

		# Translators: The label of the read only field showing the full text of the selected item.
		detailLabel = _("&Details:")
		self.detailCtrl = helper.addLabeledControl(
			detailLabel,
			wx.TextCtrl,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_DONTWRAP,
			size=(conf["dialogWidth"] - 60, 90),
		)

		optionsHelper = guiHelper.BoxSizerHelper(self, orientation=wx.HORIZONTAL)
		# Translators: The label of a combo box for filtering the list by author.
		self.authorChoice = optionsHelper.addLabeledControl(
			_("&Author:"), wx.Choice, choices=[ALL_AUTHORS]
		)
		self.authorChoice.SetSelection(0)
		self.authorChoice.Bind(wx.EVT_CHOICE, self.onFilterChanged)
		# Translators: The label of a combo box for choosing how the list is sorted.
		self.sortChoice = optionsHelper.addLabeledControl(
			_("&Sort by:"), wx.Choice, choices=self.getSortOptions()
		)
		self.sortChoice.SetSelection(0)
		self.sortChoice.Bind(wx.EVT_CHOICE, self.onFilterChanged)
		self.addExtraOptions(optionsHelper)
		helper.addItem(optionsHelper)

		self.statusText = helper.addItem(wx.StaticText(self, label=""))

		buttonHelper = guiHelper.ButtonHelper(wx.HORIZONTAL)
		# Translators: The label of the button that moves the cursor to the selected item.
		self.goToButton = buttonHelper.addButton(self, label=_("&Go to"))
		self.goToButton.Bind(wx.EVT_BUTTON, self.onGoTo)
		self.goToButton.SetDefault()
		self.addExtraButtons(buttonHelper)
		# Translators: The label of the button that copies the selected item to the clipboard.
		self.copyButton = buttonHelper.addButton(self, label=_("&Copy"))
		self.copyButton.Bind(wx.EVT_BUTTON, self.onCopy)
		# Translators: The label of the button that copies every listed item to the clipboard.
		self.copyAllButton = buttonHelper.addButton(self, label=_("Copy a&ll"))
		self.copyAllButton.Bind(wx.EVT_BUTTON, self.onCopyAll)
		# Translators: The label of the button that closes the dialog.
		buttonHelper.addButton(self, wx.ID_CANCEL, label=_("Cl&ose"))
		helper.addDialogDismissButtons(buttonHelper)
		# wx dismisses a modal dialog on a wx.ID_CANCEL button by itself, so this
		# only needs to make Escape do the same thing.
		self.SetEscapeId(wx.ID_CANCEL)

		mainSizer.Add(helper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL | wx.EXPAND)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)
		self.SetSize((conf["dialogWidth"], conf["dialogHeight"]))
		self.CentreOnScreen()

		self.populateAuthors()
		self.refreshList()
		self.filterCtrl.SetFocus()

	# -- Hooks for subclasses -------------------------------------------------

	def getListLabel(self) -> str:
		raise NotImplementedError

	def getColumns(self):
		raise NotImplementedError

	def getRowValues(self, item) -> List[str]:
		raise NotImplementedError

	def getDetailText(self, item) -> str:
		raise NotImplementedError

	def getClipboardText(self, item) -> str:
		return self.getDetailText(item)

	def getSortOptions(self) -> List[str]:
		return [
			# Translators: A sort order for the item list.
			_("Document order"),
			# Translators: A sort order for the item list.
			_("Author"),
			# Translators: A sort order for the item list.
			_("Date"),
		]

	def sortItems(self, items):
		mode = self.sortChoice.GetSelection()
		if mode == 1:
			return sorted(items, key=lambda i: ((i.author or "").lower(), i.index))
		if mode == 2:
			return sorted(items, key=lambda i: (i.date or datetime.datetime.min, i.index))
		return items

	def addExtraOptions(self, helper):
		pass

	def addExtraButtons(self, buttonHelper):
		pass

	def passesExtraFilter(self, item) -> bool:
		return True

	def getStatusText(self) -> str:
		# Translators: Reports how many items are shown out of the total.
		return _("Showing {shown} of {total}").format(
			shown=len(self.filteredItems), total=len(self.allItems)
		)

	def reload(self):
		"""Re-read the items from Word after the document was changed."""
		raise NotImplementedError

	# -- List handling --------------------------------------------------------

	def populateAuthors(self):
		authors = sorted({(i.author or "").strip() for i in self.allItems if (i.author or "").strip()})
		current = self.authorChoice.GetStringSelection()
		self.authorChoice.Set([ALL_AUTHORS] + authors)
		if current and current in authors:
			self.authorChoice.SetStringSelection(current)
		else:
			self.authorChoice.SetSelection(0)

	def matchesFilter(self, item, needle: str) -> bool:
		if not needle:
			return True
		haystack = " ".join(
			str(part or "") for part in (item.author, item.text, getattr(item, "anchorText", ""))
		).lower()
		# Every word must appear somewhere, which makes multi word filters useful.
		return all(word in haystack for word in needle.split())

	def refreshList(self, keepIndex: Optional[int] = None):
		needle = self.filterCtrl.GetValue().strip().lower()
		author = self.authorChoice.GetStringSelection()
		items = [
			i
			for i in self.allItems
			if self.passesExtraFilter(i)
			and (author == ALL_AUTHORS or (i.author or "").strip() == author)
			and self.matchesFilter(i, needle)
		]
		self.filteredItems = self.sortItems(items)

		self._suppressDetailUpdate = True
		self.itemsList.DeleteAllItems()
		for row, item in enumerate(self.filteredItems):
			values = self.getRowValues(item)
			self.itemsList.InsertItem(row, values[0])
			for col, value in enumerate(values[1:], start=1):
				self.itemsList.SetItem(row, col, value)
		self._suppressDetailUpdate = False

		self.statusText.SetLabel(self.getStatusText())
		hasItems = bool(self.filteredItems)
		for button in (self.goToButton, self.copyButton, self.copyAllButton):
			button.Enable(hasItems)
		self.updateButtonStates()
		if hasItems:
			target = 0
			if keepIndex is not None:
				target = max(0, min(keepIndex, len(self.filteredItems) - 1))
			self.itemsList.Select(target)
			self.itemsList.Focus(target)
			self.updateDetail()
		else:
			self.detailCtrl.SetValue("")

	def updateButtonStates(self):
		pass

	def getPageColumn(self) -> Optional[int]:
		"""Index of the Page column, or None if this list has none."""
		return None

	def getSelectedItem(self):
		row = self.itemsList.GetFirstSelected()
		if row == -1 or row >= len(self.filteredItems):
			return None
		return self.filteredItems[row]

	def fetchPageForSelection(self):
		"""Fill in the page number for the row the user is on.

		Page numbers make Word paginate, so they are read one row at a time
		instead of for the whole document up front, and cached afterwards.
		"""
		column = self.getPageColumn()
		if column is None or not getConf()["reportPageNumbers"]:
			return
		row = self.itemsList.GetFirstSelected()
		item = self.getSelectedItem()
		if item is None or item.page is not None:
			return
		page = self.backend.getPage(item)
		if page is not None and 0 <= row < self.itemsList.GetItemCount():
			self.itemsList.SetItem(row, column, str(page))

	def updateDetail(self):
		item = self.getSelectedItem()
		if item is not None:
			self.fetchPageForSelection()
		self.detailCtrl.SetValue(self.getDetailText(item) if item else "")

	# -- Events ---------------------------------------------------------------

	def onFilterChanged(self, evt):
		self.refreshList()

	def onSelectionChanged(self, evt):
		evt.Skip()
		if self._suppressDetailUpdate:
			return
		self.updateDetail()
		self.updateButtonStates()

	def onFilterKeyDown(self, evt):
		key = evt.GetKeyCode()
		if key in (wx.WXK_DOWN, wx.WXK_UP) and self.filteredItems:
			# Arrowing out of the filter should land in the list, as in a browser
			# address bar. This is the fastest path: type, arrow down, Enter.
			self.itemsList.SetFocus()
			return
		if key == wx.WXK_RETURN and self.filteredItems:
			self.activateSelection()
			return
		evt.Skip()

	def onListKeyDown(self, evt):
		key = evt.GetKeyCode()
		if key == wx.WXK_DELETE:
			self.onDeleteKey()
			return
		if key == wx.WXK_BACK:
			# Backspace in the list returns to the filter and erases a character,
			# so refining a search never needs a Tab press.
			self.filterCtrl.SetFocus()
			value = self.filterCtrl.GetValue()
			self.filterCtrl.SetValue(value[:-1])
			self.filterCtrl.SetInsertionPointEnd()
			return
		if key == wx.WXK_F2:
			self.onEditKey()
			return
		evt.Skip()

	def onDeleteKey(self):
		pass

	def onEditKey(self):
		pass

	def onItemActivated(self, evt):
		self.activateSelection()

	def activateSelection(self):
		item = self.getSelectedItem()
		if not item:
			return
		self.chosenItem = item
		self.saveState()
		self.EndModal(wx.ID_OK)

	def onGoTo(self, evt):
		self.activateSelection()

	def onCopy(self, evt):
		item = self.getSelectedItem()
		if item:
			self.copyToClipboard(self.getClipboardText(item))

	def onCopyAll(self, evt):
		if not self.filteredItems:
			return
		text = "\n\n".join(self.getClipboardText(i) for i in self.filteredItems)
		self.copyToClipboard(text)

	def copyToClipboard(self, text: str):
		import api as nvdaApi
		import ui

		if nvdaApi.copyToClip(text):
			# Translators: Reported when the selected items were copied to the clipboard.
			ui.message(_("Copied to clipboard"))
		else:
			# Translators: Reported when copying to the clipboard failed.
			ui.message(_("Could not copy to clipboard"))

	def saveState(self):
		conf = getConf()
		if conf["rememberFilter"]:
			conf["lastFilterText"] = self.filterCtrl.GetValue()
		# Clamp to the range the config spec allows; on a very large screen a
		# maximised dialog can otherwise exceed it and fail validation on write.
		size = self.GetSize()
		conf["dialogWidth"] = max(400, min(int(size.width), 4000))
		conf["dialogHeight"] = max(300, min(int(size.height), 3000))

	def reloadFromWord(self, keepIndex: Optional[int] = None):
		"""Re-read from Word and rebuild the list, keeping roughly the same spot."""
		try:
			self.allItems = self.reload()
		except OfficeAccessError as e:
			gui.messageBox(str(e), _("Comment Commander"), wx.OK | wx.ICON_ERROR, self)
			return
		self.populateAuthors()
		self.refreshList(keepIndex=keepIndex)


class CommentsDialog(BaseItemsDialog):
	"""The list of comments."""

	def __init__(self, parent, backend, doc, items):
		self.showResolvedCheckBox = None
		self.kindChoice = None
		# Translators: The title of the comments list dialog. {name} is the document name.
		title = _("Comments in {name}").format(
			name=backend.getDocumentName(doc) or _("this document")
		)
		super().__init__(parent, backend, doc, items, title)

	def getListLabel(self):
		# Translators: The label of the list of comments.
		return _("Co&mments:")

	def getColumns(self):
		if self.isExcel:
			return [
				# Translators: A column header in the comments list.
				(_("#"), 50),
				# Translators: A column header in the Excel comments list, the worksheet name.
				(_("Sheet"), 110),
				# Translators: A column header in the Excel comments list, the cell reference.
				(_("Cell"), 60),
				# Translators: A column header in the Excel comments list, note or comment thread.
				(_("Type"), 110),
				# Translators: A column header in the comments list.
				(_("Author"), 130),
				# Translators: A column header in the comments list.
				(_("Comment"), 260),
				# Translators: A column header in the Excel comments list, the contents of the cell.
				(_("Cell text"), 160),
				# Translators: A column header in the comments list.
				(_("Replies"), 60),
				# Translators: A column header in the comments list.
				(_("Date"), 130),
			]
		return [
			# Translators: A column header in the comments list.
			(_("#"), 50),
			# Translators: A column header in the comments list.
			(_("Author"), 130),
			# Translators: A column header in the comments list.
			(_("Comment"), 280),
			# Translators: A column header in the comments list.
			(_("Annotated text"), 200),
			# Translators: A column header in the comments list.
			(_("Status"), 80),
			# Translators: A column header in the comments list.
			(_("Replies"), 60),
			# Translators: A column header in the comments list.
			(_("Page"), 55),
			# Translators: A column header in the comments list.
			(_("Date"), 130),
		]

	def getRowValues(self, item):
		if item.isReply:
			# Translators: Marks a reply in the comments list. {number} is the reply number.
			number = _("reply {number}").format(number=item.threadPosition)
			comment = f"    {item.text}"
		else:
			number = str(item.index)
			comment = item.text
		if self.isExcel:
			return [
				number,
				item.sheetName,
				item.cellAddress,
				item.kindText,
				item.author,
				comment[:COLUMN_TEXT_LIMIT],
				item.anchorText[:COLUMN_TEXT_LIMIT],
				str(item.replyCount) if item.replyCount else "",
				formatDate(item.date),
			]
		return [
			number,
			item.author,
			comment[:COLUMN_TEXT_LIMIT],
			item.anchorText[:COLUMN_TEXT_LIMIT],
			item.statusText,
			str(item.replyCount) if item.replyCount else "",
			str(item.page) if item.page else "",
			formatDate(item.date),
		]

	def getDetailText(self, item):
		lines = []
		if self.isExcel:
			if item.sheetName:
				# Translators: A field label in the details of an Excel comment.
				lines.append(_("Sheet: {sheet}").format(sheet=item.sheetName))
			if item.cellAddress:
				# Translators: A field label in the details of an Excel comment.
				lines.append(_("Cell: {cell}").format(cell=item.cellAddress))
			if item.kindText:
				# Translators: A field label in the details of a comment.
				lines.append(_("Type: {type}").format(type=item.kindText))
		if item.author:
			# Translators: A field label in the details of a comment.
			lines.append(_("Author: {author}").format(author=item.author))
		if item.date:
			# Translators: A field label in the details of a comment.
			lines.append(_("Date: {date}").format(date=formatDate(item.date)))
		if item.page:
			# Translators: A field label in the details of a comment.
			lines.append(_("Page: {page}").format(page=item.page))
		if item.done is not None:
			# Translators: A field label in the details of a comment.
			lines.append(_("Status: {status}").format(status=item.statusText))
		if item.anchorText:
			if self.isExcel:
				# Translators: A field label in the details of an Excel comment.
				lines.append(_("Cell text: {text}").format(text=item.anchorText))
			else:
				# Translators: A field label in the details of a comment.
				lines.append(_("Annotated text: {text}").format(text=item.anchorText))
		# Translators: A field label in the details of a comment.
		lines.append(_("Comment: {text}").format(text=item.text))
		return "\n".join(lines)

	def getSortOptions(self):
		if self.isExcel:
			return super().getSortOptions() + [
				# Translators: A sort order for the Excel comments list.
				_("Sheet and cell")
			]
		return super().getSortOptions() + [
			# Translators: A sort order for the comments list.
			_("Status")
		]

	def sortItems(self, items):
		if self.sortChoice.GetSelection() == 3:
			if self.isExcel:
				return sorted(items, key=lambda i: i.sortKey)
			return sorted(items, key=lambda i: (bool(i.done), i.index))
		return super().sortItems(items)

	def addExtraOptions(self, helper):
		if self.isExcel:
			# Excel has no resolved state, but it does have two kinds of
			# annotation, so offer that as the extra filter instead.
			self.kindChoice = helper.addLabeledControl(
				# Translators: The label of a combo box filtering Excel notes and comments.
				_("&Kind:"),
				wx.Choice,
				choices=[
					# Translators: A choice in the Excel kind filter, showing everything.
					_("All kinds"),
					# Translators: A choice in the Excel kind filter, the classic yellow notes.
					_("Notes"),
					# Translators: A choice in the Excel kind filter, modern threaded comments.
					_("Comment threads"),
				],
			)
			self.kindChoice.SetSelection(0)
			self.kindChoice.Bind(wx.EVT_CHOICE, self.onFilterChanged)
			return
		# Translators: The label of a check box that shows or hides resolved comments.
		label = _("Show &resolved")
		self.showResolvedCheckBox = helper.addItem(wx.CheckBox(self, label=label))
		self.showResolvedCheckBox.SetValue(getConf()["showResolvedByDefault"])
		self.showResolvedCheckBox.Bind(wx.EVT_CHECKBOX, self.onFilterChanged)

	def passesExtraFilter(self, item):
		if self.isExcel:
			if self.kindChoice is None:
				return True
			selection = self.kindChoice.GetSelection()
			if selection == 1:
				return item.kind == KIND_NOTE
			if selection == 2:
				return item.kind == KIND_THREADED
			return True
		if self.showResolvedCheckBox is None or self.showResolvedCheckBox.GetValue():
			return True
		return not item.done

	def addExtraButtons(self, buttonHelper):
		# Translators: The label of the button that adds a reply to the selected comment.
		self.replyButton = buttonHelper.addButton(self, label=_("&Reply..."))
		self.replyButton.Bind(wx.EVT_BUTTON, self.onReply)
		# Translators: The label of the button that edits the text of the selected comment.
		self.editButton = buttonHelper.addButton(self, label=_("&Edit..."))
		self.editButton.Bind(wx.EVT_BUTTON, self.onEdit)
		# Translators: The label of the button that resolves or reopens the selected comment.
		self.resolveButton = buttonHelper.addButton(self, label=_("Resol&ve"))
		self.resolveButton.Bind(wx.EVT_BUTTON, self.onToggleResolved)
		# Translators: The label of the button that deletes the selected comment.
		self.deleteButton = buttonHelper.addButton(self, label=_("&Delete"))
		self.deleteButton.Bind(wx.EVT_BUTTON, self.onDelete)

	def updateButtonStates(self):
		item = self.getSelectedItem()
		hasItem = item is not None
		for button in (self.editButton, self.deleteButton):
			button.Enable(hasItem)
		# An Excel note is a single block of text with no thread behind it.
		self.replyButton.Enable(hasItem and item.supportsReplies)
		canResolve = hasItem and item.supportsResolve
		self.resolveButton.Enable(canResolve)
		if canResolve and item.done:
			# Translators: The label of the button that reopens a resolved comment.
			self.resolveButton.SetLabel(_("Reopen"))
		else:
			# Translators: The label of the button that resolves the selected comment.
			self.resolveButton.SetLabel(_("Resol&ve"))

	def getStatusText(self):
		threads = [i for i in self.allItems if not i.isReply]
		shown = len(self.filteredItems)
		if self.isExcel:
			notes = len([i for i in threads if i.kind == KIND_NOTE])
			# Translators: Reports the state of the Excel comments list.
			return _("{shown} shown, {notes} notes, {threads} comment threads").format(
				shown=shown, notes=notes, threads=len(threads) - notes
			)
		resolved = len([i for i in threads if i.done])
		# Translators: Reports the state of the comments list.
		return _("{shown} shown, {total} comments, {resolved} resolved").format(
			shown=shown, total=len(threads), resolved=resolved
		)

	def getPageColumn(self):
		# Excel already knows where every item lives, so nothing is fetched lazily.
		return None if self.isExcel else 6

	def reload(self):
		return self.backend.getComments(self.doc)

	# -- Actions --------------------------------------------------------------

	def _currentRow(self):
		return self.itemsList.GetFirstSelected()

	def onReply(self, evt):
		item = self.getSelectedItem()
		if not item:
			return
		# Translators: The message in the dialog for typing a reply to a comment.
		message = _("Reply to the comment by {author}:").format(author=item.author or "")
		# Translators: The title of the dialog for typing a reply to a comment.
		with wx.TextEntryDialog(self, message, _("Reply"), style=wx.TE_MULTILINE | wx.OK | wx.CANCEL) as dlg:
			if dlg.ShowModal() != wx.ID_OK:
				return
			text = dlg.GetValue().strip()
		if not text:
			return
		row = self._currentRow()
		try:
			self.backend.addReply(item, text)
		except OfficeAccessError as e:
			gui.messageBox(str(e), _("Comment Commander"), wx.OK | wx.ICON_ERROR, self)
			return
		self.reloadFromWord(keepIndex=row)

	def onEdit(self, evt):
		item = self.getSelectedItem()
		if not item:
			return
		# Translators: The message in the dialog for editing the text of a comment.
		message = _("Comment text:")
		# Translators: The title of the dialog for editing the text of a comment.
		with wx.TextEntryDialog(
			self, message, _("Edit comment"), value=item.text, style=wx.TE_MULTILINE | wx.OK | wx.CANCEL
		) as dlg:
			if dlg.ShowModal() != wx.ID_OK:
				return
			text = dlg.GetValue()
		row = self._currentRow()
		try:
			self.backend.setCommentText(item, text)
		except OfficeAccessError as e:
			gui.messageBox(str(e), _("Comment Commander"), wx.OK | wx.ICON_ERROR, self)
			return
		self.reloadFromWord(keepIndex=row)

	def onToggleResolved(self, evt):
		item = self.getSelectedItem()
		if not item or item.done is None:
			return
		row = self._currentRow()
		try:
			self.backend.setCommentDone(item, not item.done)
		except OfficeAccessError as e:
			gui.messageBox(str(e), _("Comment Commander"), wx.OK | wx.ICON_ERROR, self)
			return
		self.reloadFromWord(keepIndex=row)

	def onDelete(self, evt):
		item = self.getSelectedItem()
		if not item:
			return
		if getConf()["confirmDelete"]:
			# Translators: The confirmation message shown before deleting a comment.
			message = _("Delete the comment by {author}?\n\n{text}").format(
				author=item.author or "", text=item.text
			)
			# Translators: The title of the confirmation dialog shown before deleting a comment.
			if gui.messageBox(message, _("Delete comment"), wx.YES | wx.NO | wx.ICON_WARNING, self) != wx.YES:
				return
		row = self._currentRow()
		try:
			self.backend.deleteItem(item)
		except OfficeAccessError as e:
			gui.messageBox(str(e), _("Comment Commander"), wx.OK | wx.ICON_ERROR, self)
			return
		self.reloadFromWord(keepIndex=row)

	def onDeleteKey(self):
		self.onDelete(None)

	def onEditKey(self):
		self.onEdit(None)


class RevisionsDialog(BaseItemsDialog):
	"""The list of tracked changes."""

	def __init__(self, parent, backend, doc, items):
		# Translators: The title of the tracked changes list dialog. {name} is the document name.
		title = _("Tracked changes in {name}").format(
			name=backend.getDocumentName(doc) or _("this document")
		)
		super().__init__(parent, backend, doc, items, title)

	def getListLabel(self):
		# Translators: The label of the list of tracked changes.
		return _("Tracked &changes:")

	def getColumns(self):
		return [
			# Translators: A column header in the tracked changes list.
			(_("#"), 50),
			# Translators: A column header in the tracked changes list.
			(_("Type"), 130),
			# Translators: A column header in the tracked changes list.
			(_("Author"), 130),
			# Translators: A column header in the tracked changes list.
			(_("Text"), 330),
			# Translators: A column header in the tracked changes list.
			(_("Page"), 55),
			# Translators: A column header in the tracked changes list.
			(_("Date"), 130),
		]

	def getRowValues(self, item):
		text = item.text or item.formatDescription
		return [
			str(item.index),
			item.typeText,
			item.author,
			text[:COLUMN_TEXT_LIMIT],
			str(item.page) if item.page else "",
			formatDate(item.date),
		]

	def getDetailText(self, item):
		lines = []
		# Translators: A field label in the details of a tracked change.
		lines.append(_("Type: {type}").format(type=item.typeText))
		if item.author:
			# Translators: A field label in the details of a tracked change.
			lines.append(_("Author: {author}").format(author=item.author))
		if item.date:
			# Translators: A field label in the details of a tracked change.
			lines.append(_("Date: {date}").format(date=formatDate(item.date)))
		if item.page:
			# Translators: A field label in the details of a tracked change.
			lines.append(_("Page: {page}").format(page=item.page))
		if item.formatDescription:
			# Translators: A field label in the details of a tracked change.
			lines.append(_("Formatting: {description}").format(description=item.formatDescription))
		if item.text:
			# Translators: A field label in the details of a tracked change.
			lines.append(_("Text: {text}").format(text=item.text))
		return "\n".join(lines)

	def getSortOptions(self):
		return super().getSortOptions() + [
			# Translators: A sort order for the tracked changes list.
			_("Type")
		]

	def sortItems(self, items):
		if self.sortChoice.GetSelection() == 3:
			return sorted(items, key=lambda i: (i.typeText.lower(), i.index))
		return super().sortItems(items)

	def addExtraButtons(self, buttonHelper):
		# Translators: The label of the button that accepts the selected tracked change.
		self.acceptButton = buttonHelper.addButton(self, label=_("&Accept"))
		self.acceptButton.Bind(wx.EVT_BUTTON, self.onAccept)
		# Translators: The label of the button that rejects the selected tracked change.
		self.rejectButton = buttonHelper.addButton(self, label=_("Re&ject"))
		self.rejectButton.Bind(wx.EVT_BUTTON, self.onReject)

	def updateButtonStates(self):
		hasItem = self.getSelectedItem() is not None
		self.acceptButton.Enable(hasItem)
		self.rejectButton.Enable(hasItem)

	def getPageColumn(self):
		return 4

	def reload(self):
		return self.backend.getRevisions(self.doc)

	def _apply(self, action):
		item = self.getSelectedItem()
		if not item:
			return
		row = self.itemsList.GetFirstSelected()
		try:
			action(item)
		except OfficeAccessError as e:
			gui.messageBox(str(e), _("Comment Commander"), wx.OK | wx.ICON_ERROR, self)
			return
		self.reloadFromWord(keepIndex=row)

	def onAccept(self, evt):
		self._apply(self.backend.acceptRevision)

	def onReject(self, evt):
		self._apply(self.backend.rejectRevision)

	def onDeleteKey(self):
		self.onReject(None)
