# Agent discovery and installation

One source directory contains the package, `SKILL.md`, references and scripts. `scripts/install.ps1` creates directory junctions to that directory. It does not copy or maintain separate skill versions, and it does not edit agent configuration files.

| Agent | User scope | Project scope |
| --- | --- | --- |
| Codex CLI/App | `~/.agents/skills/academic-docx-editor` | `<project>/.agents/skills/academic-docx-editor` |
| Pi Coding Agent | `~/.agents/skills/academic-docx-editor` | `<project>/.agents/skills/academic-docx-editor` |
| Claude Code | `~/.claude/skills/academic-docx-editor` | `<project>/.claude/skills/academic-docx-editor` |

Codex and Pi share the same `.agents/skills` link. Some Codex installations also use `~/.codex/skills`; the installer offers `-CodexLegacy` for that user-level compatibility path. Avoid installing both locations with the same name unless your Codex build needs it, as duplicate discovery may occur. Claude Code follows symlinked skill folders; Windows junctions implement the same directory indirection.

Official discovery sources checked on 2026-10-13:

- [Codex skills](https://developers.openai.com/codex/skills)
- [Pi skills](https://github.com/earendil-works/pi/blob/main/packages/coding-agent/docs/skills.md)
- [Claude Code skills](https://code.claude.com/docs/en/skills)
- [Agent Skills specification](https://agentskills.io/specification)
