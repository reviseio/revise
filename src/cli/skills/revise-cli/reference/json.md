# Scripting revise: --json, exit codes, error codes

Read this when a script or program runs revise and needs to branch on the
result. Full contract: https://revise.io/developer/docs/revise/json.md (also
shipped as a JSON Schema: `revise/cli-output.schema.json`).

## The envelope

With `--json`, every command prints exactly one JSON object on stdout and
nothing on stderr:

```json
{"schema":1,"ok":true,"command":"status","meta":{"version":"3.0.0","build":"…"},
 "data":{…},"warnings":[],"error":null}
```

- `ok` is true exactly when `error` is null. `data` can be present on failure
  (what was saved before the failure) and null on success with nothing to say.
- `command` is the canonical name (`xml patch`, `skill install`; `call` for
  `revise <tool> <file>`; `help` for any `--help`).
- Within schema 1, fields are only added. Ignore unknown fields; handle unknown
  codes generically.
- Inside `data`: camelCase keys, counts end in `Count`, paths are absolute,
  n/a values are `null`, empty lists are `[]`.
- Writing a document to stdout (`revise convert x.docx -`) can't combine with
  `--json`.

## Exit codes

| Code | Meaning                                                                                       |
| ---- | --------------------------------------------------------------------------------------------- |
| 0    | Done, including "nothing to change".                                                          |
| 1    | Failed, nothing saved. Human mode prints nothing on stdout.                                   |
| 2    | Partly done: `data` says what happened, `error` what didn't.                                  |
| 70   | Internal error, or `save_uncertain`: check the file; the next command on it settles the save. |

## Errors and warnings

`error`: `code` (stable, branch on it), `message` (for people and models;
wording may change), `hint` (a shell command to run next, or null),
`details` (structured, or null). Warnings have the same shape minus `hint`.

Common error codes: `usage`, `unknown_command`, `unknown_tool`,
`file_not_found`, `unsupported_format`, `import_failed`, `would_overwrite`
(add `--force`), `same_file`, `file_busy`, `file_changed` (another program saved
it; run again), `tool_error` (a tool refused: `details.toolCode`),
`unknown_search_result` / `stale_search_result` (search again),
`change_not_found`, `find_not_found` / `find_ambiguous` (xml patch),
`body_shape_changed` (use the tools for body text), `save_uncertain`,
`internal`.

Warnings worth acting on: `file_changed_outside` (re-read the blocks in
`details.changedBlockIds` before editing them), `changes_accepted` (the output
format can't hold tracked changes), `font_fallback`, `untracked_regions`
(compare).

## Saving

Commands that change a document report `file`, `saved` (did this command
write it), and `checkpoint` (where the file is now; `revise undo <file>` steps
back). Commands that write another file report `output` and `outputWritten`.
With `--out`, edits report `copiedFrom` (the original) and `file` (the copy).
Tool calls also report `changeIds` (tracked changes created), `html` (blocks
returned, with ids), and `result` (the tool's own result; snake_case keys like
`search_result_id`).
