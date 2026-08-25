# -*- coding: UTF-8 -*-
# Comment Commander: configuration and the NVDA settings panel.
# Copyright (C) 2026 Sensotec
# This file is covered by the GNU General Public License version 2.

import addonHandler
import config
import gui
import wx
from gui import guiHelper
from gui.settingsDialogs import SettingsPanel

addonHandler.initTranslation()

CONFIG_SECTION = "commentCommander"

#: What to speak once the caret has landed on a comment.
ANNOUNCE_CHOICES = ("none", "anchor", "comment", "both")

CONFIG_SPEC = {
	"announceAfterJump": 'option("none", "anchor", "comment", "both", default="anchor")',
	"selectAnnotatedText": "boolean(default=False)",
	"showResolvedByDefault": "boolean(default=True)",
	"reportPageNumbers": "boolean(default=True)",
	"confirmDelete": "boolean(default=True)",
	"rememberFilter": "boolean(default=False)",
	"lastFilterText": 'string(default="")',
	"dialogWidth": "integer(default=900, min=400, max=4000)",
	"dialogHeight": "integer(default=560, min=300, max=3000)",
}


def initConfig():
	config.conf.spec[CONFIG_SECTION] = CONFIG_SPEC


def getConf():
	return config.conf[CONFIG_SECTION]


class CommentCommanderSettingsPanel(SettingsPanel):
	# Translators: The title of the Comment Commander category in NVDA's settings dialog.
	title = _("Comment Commander")

	def makeSettings(self, settingsSizer):
		conf = getConf()
		helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

		# Translators: The label of a combo box in the Comment Commander settings.
		announceLabel = _("After moving to a comment, &announce:")
		announceOptions = [
			# Translators: A choice for what to announce after moving to a comment.
			_("Nothing extra"),
			# Translators: A choice for what to announce after moving to a comment.
			_("The annotated text"),
			# Translators: A choice for what to announce after moving to a comment.
			_("The comment"),
			# Translators: A choice for what to announce after moving to a comment.
			_("The annotated text and the comment"),
		]
		self.announceChoice = helper.addLabeledControl(
			announceLabel, wx.Choice, choices=announceOptions
		)
		try:
			self.announceChoice.SetSelection(ANNOUNCE_CHOICES.index(conf["announceAfterJump"]))
		except ValueError:
			self.announceChoice.SetSelection(1)

		# Translators: The label of a check box in the Comment Commander settings.
		selectLabel = _("&Select the annotated text when moving to a comment")
		self.selectCheckBox = helper.addItem(wx.CheckBox(self, label=selectLabel))
		self.selectCheckBox.SetValue(conf["selectAnnotatedText"])

		# Translators: The label of a check box in the Comment Commander settings.
		resolvedLabel = _("Show &resolved comments in the list by default")
		self.resolvedCheckBox = helper.addItem(wx.CheckBox(self, label=resolvedLabel))
		self.resolvedCheckBox.SetValue(conf["showResolvedByDefault"])

		# Translators: The label of a check box in the Comment Commander settings.
		pageLabel = _("Report &page numbers in the list")
		self.pageCheckBox = helper.addItem(wx.CheckBox(self, label=pageLabel))
		self.pageCheckBox.SetValue(conf["reportPageNumbers"])

		# Translators: The label of a check box in the Comment Commander settings.
		confirmLabel = _("Ask for &confirmation before deleting a comment")
		self.confirmCheckBox = helper.addItem(wx.CheckBox(self, label=confirmLabel))
		self.confirmCheckBox.SetValue(conf["confirmDelete"])

		# Translators: The label of a check box in the Comment Commander settings.
		rememberLabel = _("Re&member the filter text between sessions")
		self.rememberCheckBox = helper.addItem(wx.CheckBox(self, label=rememberLabel))
		self.rememberCheckBox.SetValue(conf["rememberFilter"])

	def onSave(self):
		conf = getConf()
		conf["announceAfterJump"] = ANNOUNCE_CHOICES[self.announceChoice.GetSelection()]
		conf["selectAnnotatedText"] = self.selectCheckBox.GetValue()
		conf["showResolvedByDefault"] = self.resolvedCheckBox.GetValue()
		conf["reportPageNumbers"] = self.pageCheckBox.GetValue()
		conf["confirmDelete"] = self.confirmCheckBox.GetValue()
		conf["rememberFilter"] = self.rememberCheckBox.GetValue()


def registerPanel():
	gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(CommentCommanderSettingsPanel)


def unregisterPanel():
	try:
		gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(CommentCommanderSettingsPanel)
	except ValueError:
		pass
