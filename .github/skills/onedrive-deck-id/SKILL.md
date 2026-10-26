---
name: onedrive-deck-id
description: Resolve and verify PowerPoint OneDrive file IDs with read-only Microsoft Graph. Use when given a public sharing link, embed URL, iframe, or OneDrive-relative path, when asked for a driveItem ID, or when preparing a deck entry for publication.
compatibility: Requires Node.js 24+, this repository's npm dependencies, network access, and GitHub CLI or an ONEDRIVE_CLIENT_ID environment value.
---

# OneDrive Deck ID

Use the bundled [lookup script](scripts/deck-id.mjs) for repeatable, Graph-verified IDs. Do not invent an ID, manually decode sharing tokens, or create another lookup helper.

This is a repository-local skill. It uses the root [package dependencies](../../../package.json) and the publisher's [shared read-only Graph reader](../../../scripts/onedrive.mjs). Run commands from the repository root.

## Workflow

1. Accept the owner's existing public sharing link, embed URL, or iframe. Do not require a new link or generate an embed URL. If they provide an exact OneDrive-relative file path instead, use `--path`.
2. Check that Node.js 24+ and the declared dependencies are available. Run `npm ci --ignore-scripts` only if dependencies are missing or changed. Do not install another authentication library.
3. Run the lookup command. Quote every link or path. Use one command for a batch so one Microsoft sign-in covers all inputs.

   ```powershell
   npm run deck-id -- "https://1drv.ms/p/..."
   npm run deck-id -- --path "Documents\Talks\My Deck.pptx"
   npm run deck-id -- "https://1drv.ms/p/first" "https://1drv.ms/p/second"
   ```

   The command accepts quoted iframe markup too. Repeat `--path` for multiple paths. Paths are relative to the OneDrive root, not local disk paths.
4. If prompted, show the owner Microsoft's device-code sign-in instructions and let them sign in through their own browser. Wait for completion. Never enter credentials for them or call an unfinished lookup verified.
5. Read the JSON result only after the command exits successfully. For parseable stdout, use `npm run --silent deck-id -- ...`. Sign-in prompts go to stderr.
6. Report each verified filename and complete returned `id`. Suggest a lowercase, hyphen-separated `deck-name`. Add string `deck-name` and `onedrive-id` properties to [decks.yml](../../../decks.yml) only when requested.

Example result shape, not an actual file ID:

```json
[
  {
    "id": "VERIFIED_GRAPH_FILE_ID",
    "name": "My Deck.pptx",
    "size": 12345
  }
]
```

Duplicate inputs that resolve to the same file appear once. IDs are opaque strings. Preserve punctuation and do not assume one account-specific ID format.

## Authentication and Safety

- The client ID comes from `ONEDRIVE_CLIENT_ID` in the environment, or from the `segunak/slides` Actions variable through GitHub CLI. It is not a secret.
- Without a supplied authentication cache, the command uses native MSAL device-code sign-in with delegated `Files.Read` access. Tokens stay in memory, and each separate invocation needs its own sign-in.
- A securely supplied `ONEDRIVE_AUTH_CACHE` environment value enables silent lookup, including in CI. The command never reads secret values from GitHub, writes credentials, or updates publishing secrets. Interactive lookup is disabled in CI.
- Never ask for passwords, access tokens, refresh tokens, or authentication caches in chat. Do not add a client secret, request write permissions, or change sharing settings to find an ID.
- Look up metadata only. Do not download or modify the PowerPoint, run the live publisher, or commit or push anything.
- One-time account-connection helpers still belong outside the repository. This skill is the reusable lookup tool, not a replacement for initial publisher setup.

## Resolution and Failures

The script follows HTTPS redirects on `1drv.ms` and `onedrive.live.com`, extracts `id` or `resid`, and verifies the candidate through Microsoft Graph. It confirms a `.pptx` filename and PowerPoint MIME type, then requests the returned ID again to check the same file version. The account's `cid` and shortened sharing token are not file IDs.

- If the public link cannot expose an ID, request the exact OneDrive-relative path and rerun with `--path`. That uses Graph's documented path lookup directly.
- A public redirect or viewer page alone is not verification. On an error, report the failure instead of returning a candidate as a confirmed ID.
- A file that changes during lookup must finish syncing before retrying.
- IDs survive ordinary moves and renames within the same drive. Copying or deleting and recreating a file creates a different identity.

If the owner prefers a browser-only lookup, open [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer), sign in to their personal account, and consent to `Files.Read`. Use `GET` with `v1.0`:

```http
GET https://graph.microsoft.com/v1.0/me/drive/root:/{folder-path}/{file-name}.pptx?$select=id,name,file,size
GET https://graph.microsoft.com/v1.0/me/drive/root/children?$select=id,name,folder,file
GET https://graph.microsoft.com/v1.0/me/drive/items/{folder-id}/children?$select=id,name,folder,file
```

Encode path components and follow `@odata.nextLink` for folder pagination. Copy the PowerPoint file's `id`, not a folder or drive ID. Signed-out Explorer shows sample data. Do not ask the owner to share its access token.

## Checks

Run the bundled [offline tests](tests/deck-id.test.mjs) through `npm test`. To inspect manual usage without signing in, run `npm run deck-id -- --help`.

Official API references are [Get driveItem](https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0), [List children](https://learn.microsoft.com/graph/api/driveitem-list-children?view=graph-rest-1.0), and [OneDrive addressing](https://learn.microsoft.com/graph/onedrive-addressing-driveitems).
