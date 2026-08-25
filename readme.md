# Comment Commander

An NVDA add-on that makes working with comments in Microsoft Word and Excel fast.

Press `NVDA+shift+;` to get a searchable list of every comment, then press `Enter`
to jump straight to where it belongs. Reply, resolve, edit and delete without ever
opening the comments pane. In Excel both classic notes and modern comment threads
are listed together. Word tracked changes get the same treatment.

See [addon/doc/en/readme.html](addon/doc/en/readme.html) for the user
documentation (Dutch: [addon/doc/nl/readme.html](addon/doc/nl/readme.html)).

## Building

No dependencies beyond Python 3.8+ — no SCons, no add-on template.

```bash
python build.py
```

This writes the manifest, compiles the translations and produces
`commentCommander-<version>.nvda-addon`. Press `Enter` on that file to install it,
or use NVDA's Add-on Store.

To remove build output:

```bash
python build.py clean
```

## Layout

| Path | Purpose |
| --- | --- |
| `buildVars.py` | Version, author and other metadata; the manifest is generated from it. |
| `build.py` | Manifest generation, `.po` → `.mo` compilation, packaging. |
| `addon/globalPlugins/commentCommander/common.py` | Data classes and COM plumbing shared by both backends. |
| `addon/globalPlugins/commentCommander/wordAccess.py` | Microsoft Word backend. |
| `addon/globalPlugins/commentCommander/excelAccess.py` | Microsoft Excel backend. |
| `addon/globalPlugins/commentCommander/backend.py` | Picks the backend that matches the focused application. |
| `addon/globalPlugins/commentCommander/dialogs.py` | The list dialogs. |
| `addon/globalPlugins/commentCommander/settings.py` | Config spec and the NVDA settings panel. |
| `addon/globalPlugins/commentCommander/__init__.py` | The global plugin: scripts and gestures. |
| `addon/locale/<lang>/LC_MESSAGES/nvda.po` | Translations. Compiled to `.mo` at build time. |
| `addon/doc/<lang>/readme.html` | User documentation shown by NVDA's add-on help. |

## Design notes

**One data class, two applications.** Both backends expose the same functions and
return the same `CommentItem`, so the dialogs never branch on which application
supplied the data - only on which columns to show. Adding a third application
means adding a module and one entry in `backend.BACKENDS`.

**Excel has two kinds of annotation.** Classic notes come from
`Worksheet.Comments`, modern threads from `Worksheet.CommentsThreaded`. Microsoft's
documentation is contradictory about whether the threaded collection also returns
legacy notes, so both are read and then deduplicated by cell address: correct under
either reading. Notes have no date, no replies and no resolved state; threads have
a date and replies but Excel exposes no resolved state either.

**Two COM calls that must never happen by accident.** comtypes' dynamic dispatch
tries a property get before treating a name as a method. For Excel's
`Text(Text, Start, Overwrite)` and `AddCommentThreaded(Text)` every argument is
optional, so a bare attribute lookup *runs* them - which for `Text` deletes the
comment body and for `AddCommentThreaded` creates an empty comment. Reads go
through `common.readValue`, which refuses to invoke a callable, and writes go
through `common.flagAsMethod` first.

**Why a global plugin rather than an app module.** An add-on app module for
`winword` would have to subclass NVDA's built-in one via `nvdaBuiltin`, and only
one add-on can win that slot. A global plugin that checks for Word on each command
cannot collide with other add-ons.

**Finding the document.** Each backend tries NVDA's own object model handle on the
focus ancestry first, then `AccessibleObjectFromWindow` with `OBJID_NATIVEOM` on
the window handles around focus. Both routes start from the focused window, so
neither can reach a document belonging to a background application. An earlier
version also consulted the running object table, which returned Word's active
document from anywhere - that is why it is gone.

**No keyword arguments over COM.** NVDA ships comtypes' dynamic dispatch, whose
`MethodCaller.__call__` takes positional arguments only. Every Word call here is
positional — hence `Replies.Add(range, text)` rather than `Add(Text:=...)`.

**Jump timing.** Enter in the dialog records the choice and closes; the caret is
only moved after `postPopup()` has handed focus back to Word, on a short
`wx.CallLater`. Moving it while the dialog still had focus made Word's caret land
inconsistently.

## Localisation

`addon/locale/nl/LC_MESSAGES/nvda.po` holds the Dutch translation of all 143
strings. To add a language, copy that directory, translate the `msgstr` entries and
rebuild. Fuzzy and empty entries are skipped at compile time and fall back to
English.

## Support and how this was built

This add-on was built with the help of AI. It has been tested thoroughly, but I
cannot guarantee support for it.

Bug reports and pull requests are welcome — please open an issue. Treat a reply or
a fix as a favour rather than a promise.

To be precise about what "tested" means here: the add-on is covered by an
extensive automated test suite that exercises the logic against a simulated Word
and Excel object model. That catches a great deal, but it is not the same as
years of use in the field. Try a new version on a copy of an important document
before trusting it with one that matters.

## Licence

GNU General Public License version 2.
