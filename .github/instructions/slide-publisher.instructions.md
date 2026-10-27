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
- Download source PowerPoints only into the runner's temporary directory, outside the checkout. Never commit the originals or include them in Pages output. Do not modify OneDrive sources.
- Run the pinned PPTX2HTML CLI with its interactive player, separate assets, and strict conversion. Preserve notes, hidden slides, animations, transitions, and the converter's normal functionality.
- Preserve authored text formatting through animation, backward navigation, and replay. Guard local source corrections against incompatible upstream changes and leave the installed dependency unchanged.
- Test and convert on Windows. Keep installed or embedded original fonts first, then self-host needed Google Fonts fallbacks in the generated presentation. Reuse [PPTX2HTML's mappings and classification](../../node_modules/pptx2html/src/render/fonts.js), with [Fontconfig's substitution reference](https://github.com/fontconfig/fontconfig/blob/main/conf.d/30-metric-aliases.conf) as research guidance, not an installed dependency. Do not maintain a second exhaustive source-font map or use AI during publishing.
- Load bundled fonts before the player lays out slides. Verify Chromium's actual rendered fonts and log substitutions. Unknown fonts get a category fallback; missing assets or broken output still stop publishing. Keep font license notices alongside generated assets. Never copy Windows font binaries into the website.
- Verify staged presentations with Playwright headless Chromium before publishing. Export PDFs with Chromium, with one correctly sized page per slide and completed reveals. Validate page count, dimensions, final text, and loaded assets.
- Verify embedded audio/video decoding, playback, and seeking. Normal canceled media-range requests are not failures when the media checks pass; missing or unsupported media still stops publication.
- Publish unaltered originals and validated PDFs as stable assets in one rolling GitHub Release. Replace edited decks' download pairs rather than accumulating versions. Leave unrelated release assets alone. Original PowerPoints and PDFs stay out of Git and Pages.
- Keep extracted media above Git's file-size limit in that same rolling release and reference its stable URLs from the presentation. Verify locally before moving media out of Pages output. Replace media for edited decks, remove their obsolete assets, and recover missing release assets from existing presentation references without another manifest.
- `decks` contains generated HTML, media, a public deck index, and converter attribution. Never hand-edit generated presentations.
- Keep the HTML index complete and current with every successfully published deck's title, URL, and both downloads. Keep URLs selectable and easy to copy, and make title links visibly identifiable as links that open in a new tab. Presentations expose the same downloads without replacing their player controls.
- `.github/decks-state.json` only remembers each published deck's name, OneDrive ID, and last published `eTag`. The updater maintains this list to skip unchanged decks. There is no schema version or publisher fingerprint.
- After changing the converter or publishing code, use Run workflow with **Rebuild all configured decks** to regenerate unchanged source decks.
- Prepare and validate all changed output before uploading downloads or replacing the working publication. Upload failure prevents website publication. Stable downloads may briefly advance before Pages deploys, which is intentional. A missing or inaccessible source is an error, not a request to unpublish it. Remove its configuration entry to unpublish intentionally.
- Only Update Decks makes automatic commits, limited to generated output and state. If main advances during conversion, retry against its latest revision rather than force-pushing or merging stale output over source changes.
- Use path exclusions for push and pull-request triggers so new code directories are covered automatically. Documentation, repository housekeeping, and generated-only commits must not start another synchronization.
- Publish only after successful synchronization, using its exact commit confirmed on main. Code pushes and manual updates publish even without a new generated commit; scheduled runs publish only when output changed. Keep manual dispatch available on both workflows. A standalone manual Pages run publishes the latest committed output without converting decks.
- Serialize update-through-deployment with manual publishing so they cannot overlap. A reusable Pages call must not reacquire its caller's lock. Refuse output older than the presentation files already on main.
- Publish Pages only uploads the generated site and deploys it. No Jekyll, OneDrive access, or conversion belongs there. Generated commits must not retrigger the updater or independently start a second deployment.
- Keep the public-facing README free of internal setup details. Do not commit or push from an agent session.

## Local Checks

Use Node.js 24 or later. Run `npm ci --ignore-scripts` after changing dependencies, install Playwright's Chromium with `npx playwright install chromium`, then run `npm test`. Tests are offline and generate tiny original fixtures outside the repository to exercise the real converter, browser playback, PDF export, and mocked release operations. Run `npm run verify:converter` for converter changes and forced rebuilds. This runs the upstream Chromium-only checks without modifying the installed dependency.

`npm run update-decks` is the live publishing preparation command. It requires the Actions credentials, may renew the repository secret, replaces release assets, and writes generated website files and state. It does not commit or push locally. Do not run it as an offline test.
