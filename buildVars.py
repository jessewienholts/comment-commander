# -*- coding: UTF-8 -*-
"""Metadata for the Comment Commander NVDA add-on.

Edit the values here and re-run build.py; the manifest is generated from them.
"""

addon_info = {
	# Internal name; also the folder name inside NVDA's add-ons directory.
	"addon_name": "commentCommander",
	# Translators: Summary of the add-on, shown in NVDA's Add-on Store.
	"addon_summary": "Comment Commander",
	# Translators: Long description of the add-on, shown in NVDA's Add-on Store.
	"addon_description": (
		"Work with comments in Microsoft Word and Excel efficiently. "
		"Press NVDA+shift+semicolon for a searchable list of every comment, "
		"then press Enter to jump straight to it. "
		"Reply to, resolve, edit and delete comments without ever opening the comments pane. "
		"In Excel both classic notes and modern comment threads are listed together. "
		"Word tracked changes can be browsed the same way."
	),
	"addon_version": "2.0.0",
	"addon_author": "Sensotec <support@sensotec.com>",
	"addon_url": "https://www.sensotec.be",
	"addon_sourceURL": "https://github.com/jessewienholts/comment-commander",
	"addon_docFileName": "readme.html",
	"addon_minimumNVDAVersion": "2023.1",
	"addon_lastTestedNVDAVersion": "2026.1",
	"addon_updateChannel": None,
	"addon_license": "GPL v2",
	"addon_licenseURL": "https://www.gnu.org/licenses/old-licenses/gpl-2.0.html",
}

#: Python sources that carry translatable strings.
i18nSources = [
	"addon/globalPlugins/commentCommander/__init__.py",
	"addon/globalPlugins/commentCommander/backend.py",
	"addon/globalPlugins/commentCommander/common.py",
	"addon/globalPlugins/commentCommander/dialogs.py",
	"addon/globalPlugins/commentCommander/excelAccess.py",
	"addon/globalPlugins/commentCommander/settings.py",
	"addon/globalPlugins/commentCommander/wordAccess.py",
]

#: Files never included in the built add-on.
excludedFiles = []
