# GitHub Copilot Instructions

This repository publishes selected OneDrive PowerPoints as interactive HTML at `https://segunakinyemi.com/slides/<deck-name>/`. Follow the [slide publisher instructions](instructions/slide-publisher.instructions.md) when changing configuration, publishing code, or workflows. Never commit or push from an agent session.

Keep instructions focused on intent and constraints. Read schedules, versions, and other implementation details from code and configuration instead of duplicating them in agent context.

## Finding OneDrive Deck IDs

Use the [onedrive-deck-id skill](skills/onedrive-deck-id/SKILL.md) when given a public PowerPoint link, embed, or OneDrive path, or when asked to find an ID for publication. The skill owns the lookup script, manual command, authentication guidance, and verification procedure. Do not create ad hoc lookup helpers or guess IDs from sharing tokens.
