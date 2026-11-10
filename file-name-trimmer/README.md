# File Name Trimmer

A separate Windows app for trimming downloaded Word submission filenames. No Python,
Word or extra packages are required. It uses Windows PowerShell and Windows Forms.

Download this branch using **Code > Download ZIP**, extract it, open the
`file-name-trimmer` folder, and double-click **File Name Trimmer.bat**. Keep the
`.bat` and `FileNameTrimmer.ps1` together. Alternatively, download and extract the
`FileNameTrimmer-Windows` artifact from this branch's GitHub Actions build.

1. Choose a folder. Optionally enable **Include subfolders**.
2. Review the current names, proposed names and skipped-file reasons.
3. Click **Rename ready files** to rename the files in place.

Only `.docx` files are considered, including uppercase `.DOCX`. The app removes
everything up to and including the first underscore, preserving all remaining
characters, subsequent underscores, the extension, file contents and folder.

| Before | After |
| --- | --- |
| `12345_John_Bears.docx` | `John_Bears.docx` |
| `prefix_a_b_c.docx` | `a_b_c.docx` |
| `_Jane.docx` | `Jane.docx` |
| `Bears.docx` | Skipped: no underscore |

Files are skipped when the resulting name is empty or reserved by Windows, when
another file or folder already has that name, or when multiple files would receive
the same name. Word lock files (`~$...`) are skipped. Files that cannot be renamed
show an error in the preview. Existing files are never overwritten, even if a new
file appears after the preview.

Each click renames only the ready files in the displayed preview. Refreshing a
preview after renaming applies the rule again to the new names, so review it before
clicking Rename again. For example, a second pass on `John_Bears.docx` would produce
`Bears.docx`.

For automated verification on Windows:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\FileNameTrimmer.ps1 -SelfTest
```
