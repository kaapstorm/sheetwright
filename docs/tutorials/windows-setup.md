# Windows setup primer

A guided walkthrough for getting a sheetwright work environment
running on Windows 10 or 11. If you are comfortable on the command
line and just want install commands, jump to [getting
started](../getting-started.md) — it has Windows snippets alongside
the macOS and Linux ones.

This primer is for readers who do most of their work in Excel or
File Explorer and have not used PowerShell, `winget`, or git from a
terminal before. By the end you will have:

- LibreOffice installed (sheetwright's default calc engine).
- Sourcetree installed (a graphical git client).
- A project folder ready for `sheetwright init`.
- Enough PowerShell to run sheetwright commands without guessing.

## What is WinGet, and why use it?

[WinGet](https://learn.microsoft.com/en-us/windows/package-manager/winget/)
is the **Windows Package Manager**. It ships with Windows 11 and
recent Windows 10 builds as the `winget` command. Think of it as the
Windows equivalent of `apt` on Ubuntu or `brew` on macOS: a single
command that downloads, installs, and updates software from a curated
catalogue.

Why bother instead of clicking through installers?

- **Reproducible.** `winget install <id>` does the same thing on
  every machine. You can paste install commands into onboarding docs
  and trust them.
- **Updatable.** `winget upgrade --all` pulls the latest versions of
  everything you installed through it. No more hunting for "Check
  for updates" menus.
- **Scriptable.** The same commands work in a setup script, in a CI
  job, or on a freshly-imaged laptop.
- **Trustworthy sources.** Packages come from publishers' own
  installers, not random download mirrors.

To check that `winget` works, open **PowerShell** (press the Windows
key, type `powershell`, press Enter) and run:

```powershell
winget --version
```

If you see a version like `v1.7.10661`, you are ready. If not,
install **App Installer** from the Microsoft Store, then reopen
PowerShell.

## Install LibreOffice with WinGet

sheetwright uses LibreOffice's headless mode (`soffice`) to evaluate
spreadsheet formulas. Install it with:

```powershell
winget install --id TheDocumentFoundation.LibreOffice
```

WinGet will download LibreOffice, run its installer silently, and
add it to your system. To confirm:

```powershell
soffice --version
```

If PowerShell reports `soffice` is not recognised, close and reopen
PowerShell so it picks up the updated `PATH`. If it still cannot
find it, add LibreOffice's program folder to your `PATH` — typically
`C:\Program Files\LibreOffice\program` (see [PowerShell
basics](#powershell-basics-for-using-sheetwright) below).

## Install Sourcetree with WinGet

[Sourcetree](https://www.sourcetreeapp.com/) is a free graphical git
client from Atlassian. sheetwright uses git for the re-import flow
(it checks for uncommitted source changes before overwriting your
files), and Sourcetree gives you a visual way to stage, commit, and
review diffs without learning git's command line first.

```powershell
winget install --id Atlassian.Sourcetree
```

Launch Sourcetree from the Start menu. On first run it will:

1. Ask you to sign in or skip — skip is fine for local work.
2. Offer to install **Git** if it is not already on the system.
   Accept; sheetwright needs `git` available too.
3. Ask for your name and email. Use the same name and email you
   want on your commits.

You can drive sheetwright projects entirely from Sourcetree's
**Stage**, **Commit**, **Push**, and **Pull** buttons. The "Show
diff" pane is especially useful for reviewing changes to your
Markdown sheet sources before committing.

## Create a folder for your project

PowerShell uses `\` (backslash) for paths, but forward slashes work
in most commands too. Pick a location you'll remember — your user
folder is a sensible default:

```powershell
cd $HOME
mkdir my-model
cd my-model
```

`$HOME` expands to `C:\Users\<your-username>`. After the three
commands above your prompt should look something like:

```
PS C:\Users\you\my-model>
```

This is now your **project folder**. Every sheetwright command
you run should be from inside it.

If you want git tracking from the start, initialise the repo with
Sourcetree:

1. **File → Clone / New… → Create**.
2. Set **Destination Path** to `C:\Users\you\my-model`.
3. Click **Create**.

Sourcetree will turn the folder into a git repo. From here on,
every `sheetwright build` or edit will show up in Sourcetree's
**File Status** view as something you can stage and commit.

## PowerShell basics for using sheetwright

You don't need to become a PowerShell expert. The handful of
commands below cover everything sheetwright's docs assume.

### Navigating

| Task                             | PowerShell                |
|----------------------------------|---------------------------|
| Show the current folder          | `pwd` (or `Get-Location`) |
| List files in the current folder | `ls` (or `dir`)           |
| Move into a folder               | `cd path\to\folder`       |
| Move up one folder               | `cd ..`                   |
| Go to your home folder           | `cd $HOME`                |

Tab-completion works: type the first few letters of a folder or
file name and press **Tab**.

### Running sheetwright

sheetwright is invoked through `uv` (see the [getting started
guide](../getting-started.md) for installing `uv`). From inside
your project folder:

```powershell
uv run sheetwright --version
uv run sheetwright init .
uv run sheetwright build
uv run sheetwright test
```

The `.` in `init .` means "scaffold into the current folder."

### Inspecting and editing files

Open the current folder in **File Explorer**:

```powershell
explorer .
```

Open a file in your default editor (or pass a specific app):

```powershell
notepad sheetwright.toml
```

If you have **VS Code** installed via WinGet
(`winget install Microsoft.VisualStudioCode`), you can open the
project in it with:

```powershell
code .
```

### Adding a folder to PATH (only if needed)

If a command like `soffice` or `git` "is not recognised", you may
need to add its install folder to your `PATH`. The least-surprising
way is the **System Properties** GUI:

1. Press **Windows key**, type `environment variables`, press Enter.
2. Click **Environment Variables…**.
3. Under **User variables**, select **Path** and click **Edit…**.
4. Click **New**, paste the folder (e.g.
   `C:\Program Files\LibreOffice\program`), click **OK** on each
   dialog.
5. Close and reopen PowerShell.

For a one-off session you can add to `PATH` from PowerShell:

```powershell
$env:Path += ';C:\Program Files\LibreOffice\program'
```

This only affects the current PowerShell window.

### Quoting paths with spaces

If a path contains spaces, wrap it in single quotes:

```powershell
cd 'C:\Users\you\OneDrive - Acme\my-model'
```

## What's next?

- [Getting started](../getting-started.md) — install `uv`, scaffold
  a project, build, and run your first test. The Windows snippets
  there pick up where this primer leaves off.
- [Greenfield project tutorial](greenfield-project.md) — build a
  full model from scratch.
