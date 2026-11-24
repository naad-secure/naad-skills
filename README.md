# NAAD Skills

Agent skills used at NAAD, in the standard `SKILL.md` format. They work with
[OpenCode](https://opencode.ai) and Claude Code.

| Skill | What it does |
|---|---|
| [`word-live`](skills/word-live/SKILL.md) | Edits the Word document the user has open in Microsoft Word, live, while they keep working: read, fill in forms, replace text, insert images, track changes, render pages to check the layout. Built for Persian/RTL documents. Runs on Windows (agent and Word on the same PC) and on macOS with Word in a Parallels Windows VM. |

## Install

Each folder under `skills/` is one skill. Copy (or symlink) it into a folder the
agent reads skills from; the folder name must stay the same as the skill's `name`.

**macOS / Linux**

```bash
git clone git@github.com:naad-secure/naad-skills.git
mkdir -p ~/.agents/skills
ln -s "$PWD/naad-skills/skills/word-live" ~/.agents/skills/word-live
```

**Windows** (PowerShell)

```powershell
git clone git@github.com:naad-secure/naad-skills.git
New-Item -ItemType Directory -Force "$env:USERPROFILE\.agents\skills" | Out-Null
Copy-Item -Recurse naad-skills\skills\word-live "$env:USERPROFILE\.agents\skills\word-live"
```

Skill folders the agents read:

- OpenCode: `~/.agents/skills/` and `~/.claude/skills/`, and `.agents/skills/` or
  `.claude/skills/` in a project (besides its own skill folders).
- Claude Code: `~/.claude/skills/` and `.claude/skills/` in a project.

Restart the agent after installing; the skill then loads by itself when a request
matches its description (e.g. "fill in the form I have open in Word").

## word-live requirements

- **Windows:** Word open with the document active. Nothing else to install; the
  scripts use the built-in Windows PowerShell 5.1.
- **macOS:** Parallels Desktop with a running Windows VM (Parallels Tools
  installed) and Word open there. `prlctl` must be on the `PATH`.

Tested with Word on Windows 11 (Windows PowerShell 5.1), and end to end with OpenCode 2.4 on macOS.
