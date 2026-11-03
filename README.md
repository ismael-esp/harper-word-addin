# Harper Checker for Word

An offline grammar and spelling checker for Microsoft Word, built on [Harper](https://github.com/automattic/harper). It works like Grammarly's Word sidebar: colored squiggly underlines in your document, a panel with suggestions, one-click fixes, and a personal dictionary. Your text never leaves your computer.

English only.

> Unofficial project, not affiliated with or endorsed by Automattic (makers of Harper) or Microsoft.

## What's in this folder

| Folder / file | What it is |
|---|---|
| `word-addin\` | The Word add-in. This is the main thing. |
| `word-addin\Setup.cmd` | One-time setup on a new computer. |
| `word-addin\Start Harper.cmd` | Starts the checker. Must be running while you use Word. |
| `word-addin\Uninstall.cmd` | Removes the add-in from Word. |
| `word-addin\dictionary.txt` | Your personal dictionary (one word per line). Created when you add your first word; not stored in the GitHub repo. |
| `harper-app\` | A small standalone checker window (paste text, fix it, copy it back). Optional. |
| `dev\` | A fake Word test bench for developing the add-in in a normal browser (see [Development](#development)). |
| `third-party\` | Licenses of the bundled Harper code. |

## Requirements

- Windows 10 or 11
- **Word for Microsoft 365** (the desktop app)
- **Node.js** (LTS version) from <https://nodejs.org>. Install it with the default options.
- Internet during setup (it downloads a small Microsoft tool), and when the panel opens (Word loads Microsoft's add-in library). Your text is still checked locally and never sent anywhere.

## Install on a new computer (about 5 minutes)

1. **Install Node.js** from <https://nodejs.org> if it isn't installed yet.
2. **Get the files and extract them to a permanent place**, for example `Documents\Harper Checker`. From GitHub: **Code → Download ZIP**.
   Don't run it from inside the zip or from Downloads: Word remembers the folder's location.
3. **Close Word.**
4. Open the `word-addin` folder and double-click **`Setup.cmd`**.
   - Windows shows a **Security Warning** asking to install a certificate for "localhost". Click **Yes**. (Word only loads add-ins over HTTPS; this lets your PC serve the add-in to itself.)
   - The window ends with "Done". Press any key to close it.
5. Double-click **`Start Harper.cmd`**. A window says *"Harper for Word is running"*. Leave it open (you can minimize it).
6. Open Word and a document. Go to **Home → Add-ins**, and under **Developer Add-ins** click **Harper Checker**.
   After the first time, a **Check Grammar** button appears at the right end of the **Home** tab.

The first load takes a few seconds.

## Every day

1. Double-click **`Start Harper.cmd`** (keep its window open).
2. In Word, click **Check Grammar** on the Home tab.

**Start it automatically with Windows, with no window (optional):** double-click **`Install Autostart.cmd`** once. Harper then runs hidden in the background every time you log in, so you never need `Start Harper.cmd`. To stop it, end `node.exe` in Task Manager; to undo the auto-start, run `Uninstall.cmd`.

## Using the panel

- **Underlines** in your document: red = spelling, blue = grammar, purple = style.
- **Click an underline** → the panel jumps to that issue. **Click a card** → Word selects the text.
- **Fix buttons** replace the text and keep your formatting. The first button is the best suggestion.
- **Add to dictionary** (spelling issues): the word is never flagged again, in any document.
- **Ignore**: hides that kind of issue for that word.
- **Filter** (dropdown under the header): show only one type of issue.
- **↻** rechecks the whole document and redraws all underlines. Use it if underlines look out of date.
- **⚙ Settings**:
  - Your personal dictionary. Click **×** to remove a word.
  - All of Harper's rules, grouped and searchable. Untick a rule to stop it (for example, rules that are noisy for academic writing). **Reset** turns them all back on.
  - **Restore ignored issues**.
- **Dialect**: American, British, Australian, Canadian or Indian English.

If **Track Changes** ("Reviewing" mode) is on, every fix appears as a tracked change. Switch to **Editing** mode to apply fixes directly.

## What moves between computers

| Saved in | Moves with this folder? |
|---|---|
| Personal dictionary (`word-addin\dictionary.txt`) | ✅ Yes. Copy the folder (or this file) to keep your words in sync. |
| Rule settings, ignored issues, dialect | ❌ No. They're stored in Word's add-in storage on each computer. |

You can edit `dictionary.txt` in Notepad: one word per line. Changes apply the next time the panel opens.

## Troubleshooting

| What you see | What to do |
|---|---|
| Panel is blank or says "can't reach this page" | `Start Harper.cmd` isn't running. Start it, then close and reopen the panel. |
| Start window says "Certificate not found" | Run `Setup.cmd` again and click **Yes** on the certificate warning. |
| Setup says "Node.js is not installed" | Install Node.js LTS from <https://nodejs.org>, then run `Setup.cmd` again. |
| No **Harper Checker** under Developer Add-ins | Close Word completely, run `Setup.cmd` again, reopen Word. |
| "Couldn't add" / "Couldn't remove" a word | Close the `Start Harper` window, run `Start Harper.cmd` again, reopen the panel. |
| Panel shows an old version after an update | Close Word completely and reopen it (Word caches add-ins). |
| Start window says "Harper is already running" | It's already running in another window. Nothing to do. |

**Moved the folder?** Run `Setup.cmd` again from the new location.

## Uninstall

1. Run `word-addin\Uninstall.cmd`.
2. Optional, to remove the certificate: open a terminal and run `npx office-addin-dev-certs uninstall`.
3. Delete the folder.

## Standalone checker (optional)

`harper-app\Harper Checker.cmd` opens a small window (in Microsoft Edge) where you paste any text, fix it and copy it back. It needs nothing installed and works fully offline. Formatting isn't kept, so it's best for emails and short texts.

## Development

The panel is plain HTML and JavaScript in `word-addin\taskpane.html`; `word-addin\server.mjs` serves it to Word and stores the dictionary.

To debug without Word, run the fake Word test bench and open <http://127.0.0.1:4820/>:

```
node dev\serve.mjs
```

It shows an editable sample document with deliberate mistakes next to the panel, with counters for every call the add-in makes to Word and a "Big document" button for stress tests.

**Rebuilding `harper.bundle.js`** (only to update Harper): in `harper-app\`, run `npm install`, then
`npx esbuild entry.js --bundle --format=iife --global-name=Harper --platform=browser --external:fs --external:fs/promises --external:path --external:url --outfile=harper.bundle.js`, and copy the result into `word-addin\`.

## License and credits

- This project's own code: [MIT](LICENSE).
- Grammar checking by [Harper](https://github.com/automattic/harper) by Automattic, licensed under Apache-2.0. `harper.bundle.js` bundles harper.js and fflate unmodified; their licenses are in [third-party](third-party/README.md).

Harper is rule-based, so it catches spelling, grammar and common style problems, but not everything an AI checker would (for example "Me and my team has").
