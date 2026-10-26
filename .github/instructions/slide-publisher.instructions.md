---
applyTo: "decks.yml,package*.json,scripts/**,.github/skills/onedrive-deck-id/**,.github/workflows/**,.github/decks-state.json,decks/**"
---

# Slide Publisher

This repository publishes selected personal OneDrive PowerPoints at `https://segunakinyemi.com/slides/<deck-name>/`.

- `decks.yml` is the only manually maintained deck list. Each entry has string `deck-name` and `onedrive-id` values. Names are lowercase URL slugs. Before any download, reject the configuration if deck names are not unique or do not match `^[a-z0-9]+(-[a-z0-9]+)*$`. IDs are Microsoft Graph file IDs.
- Use native MSAL Node with delegated `Files.Read` access. No Rclone, client secret, interactive authentication in Actions, or model calls.
- Actions variables are `ONEDRIVE_CLIENT_ID` and `SLIDE_PUBLISHER_APP_CLIENT_ID`. Secrets are `ONEDRIVE_AUTH_CACHE` and `SLIDE_PUBLISHER_APP_PRIVATE_KEY`.
- The GitHub App needs Contents and Secrets read/write access to this repository. It saves renewed MSAL authentication before conversion. Credentials must never enter files, artifacts, caches, logs, or commits.
- One-time sign-in helpers belong outside the repository. If Microsoft revokes access, reconnect using a temporary local helper and replace the Actions secret.
- Use the [OneDrive deck-ID skill](../skills/onedrive-deck-id/SKILL.md) for read-only file-ID lookup and manual command guidance. Do not run the live publisher to discover an ID.
- Download source PowerPoints only into the runner's temporary directory, outside the checkout. Never commit or publish the original PowerPoints. Do not modify OneDrive sources.
- Run the pinned PPTX2HTML CLI with its interactive player, separate assets, and strict conversion. Preserve notes, hidden slides, animations, transitions, and the converter's normal functionality.
- `decks` contains generated HTML, media, a public deck index, and converter attribution. Never hand-edit generated presentations.
- Keep the HTML index complete and current with every successfully published deck's title and URL. It is the public deck directory.
- `.github/decks-state.json` only remembers each published deck's name, OneDrive ID, and last published `eTag`. The updater maintains this list to skip unchanged decks. There is no schema version or publisher fingerprint.
- After changing the converter or publishing code, use Run workflow with **Rebuild all configured decks** to regenerate unchanged source decks.
- Prepare and validate all changed output before replacing the working publication. A missing or inaccessible source is an error, not a request to unpublish it. Remove its configuration entry to unpublish intentionally.
- Only the Update Decks workflow makes automatic commits, limited to `decks` and generated state. Use the app token so its pushes trigger Pages. Do not force-push or automatically merge concurrent source changes.
- Publish Pages only uploads the `decks` directory and deploys it. No Jekyll, OneDrive access, or conversion belongs there. Generated commits must not retrigger the updater.
- Keep the public-facing README free of internal setup details. Do not commit or push from an agent session.

## Local Checks

Use Node.js 24 or later. Run `npm ci --ignore-scripts` after changing dependencies, then `npm test`. Tests are offline and generate a tiny original PowerPoint outside the repository to exercise the real pinned converter.

`npm run update-decks` is the live publishing preparation command. It requires the Actions credentials, may renew the repository secret, and writes generated website files and state. It does not commit or push locally. Do not run it as an offline test.
