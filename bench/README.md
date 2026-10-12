# revise benchmark

This benchmark measures how coding agents handle Word (.docx) tasks with and without the `revise` CLI. Each task runs twice, once with `revise` available and once without it. A script then grades both results.

Everything here is open: the harness, the prompts, the graders and the fixtures. You can re-run any published result on your own machine or in GitHub Actions. Each run records the exact agent version, model, `revise` version and toolchain it used.


> **Run it in CI or a disposable VM, not on your own computer.** Agents get a
> shell. The harness puts stubs for `open` and `osascript` on PATH, but an
> agent can call `/usr/bin/osascript` directly, and on a Mac that drives any
> app the terminal may automate. In one local run, a baseline agent whose
> LibreOffice conversion hung scripted the user's open Microsoft Word, sent it
> "close every document saving no", and force-killed LibreOffice. The GitHub
> Actions workflow runs on Linux runners, where there is no Word or AppleScript.

## What is measured

The benchmark runs a headless coding agent on ten document tasks. Each task runs in a fresh directory that holds copies of its input files. There are two **arms**:

| Arm | What the agent has |
|---|---|
| `revise` | The `revise` binary on `PATH` and the `revise` skill installed for the agent. |
| `baseline` | No `revise`. The agent has what a typical developer machine has for Word files: Python with `python-docx` and `lxml`, `pandoc`, headless LibreOffice, and shell tools. It also keeps the agent's own bundled skills. For Hermes that includes its `docx` skill. |

Both arms get the same prompt, model, tools, timeout and machine. Each run produces these numbers:

- **pass**: every scripted check passed. Fidelity checks are included (see Grading).
- **wall time**: seconds from launching the agent until it exits.
- **turns**:
  - Claude Code: the turn count it reports.
  - Codex: tool steps plus the final answer.
  - Hermes: API calls.
- **cost**: the agent's own cost report. For Codex, cost is computed from token usage at the model's list price. It is empty when the agent bills a subscription and reports no price.
- **tokens**: recorded where the agent reports them.

`summarize.py` reports, for each category and task, pass counts and the medians of the other numbers for each arm. It also shows the ratio baseline / revise.

## Tasks

| Category | Task | Input | Asked for |
|---|---|---|---|
| Editing | `redline3` | contract | Three clause changes (9.2, 4.1, 7.2) as tracked changes by "Claude", with a comment on each and nothing else changed |
| Editing | `bulk-rename` | contract | Every "Provider" becomes "Northstar", as tracked changes, nothing else changed |
| Editing | `long-doc-fix` | handbook (~115 pages) | Change the leave accrual to 2 days per month and the annual total to 24 days wherever it is stated, as tracked changes. A nearby "18 days' written notice" must stay untouched. |
| Editing | `table-row` | contract | Insert a table row above "Year-one total" and update the total, as tracked changes |
| Editing | `heading-fix` | handbook | Apply Heading 1 to the one section heading that was typed as bold body text, with no other change |
| Review | `compare` | contract, contract v2 | Build a redline of v2 against v1, authored by "Northstar Legal" |
| Review | `resolve-tracked` | contract with tracked changes | Accept every change except those in section 8, which are rejected |
| Review | `comment-review` | NDA | Leave at least three anchored comments by "Claude" without changing the text (see Fixtures: not included by default) |
| Create & convert | `to-pdf` | contract | Convert to PDF, keeping layout, header and footer |
| Create & convert | `generate-memo` | none | Create a memo from scratch with a title, headings, bullets, a table, a "Page X of Y" footer, US Letter paper and 1-inch margins |

The exact prompts are in `tasks.py`, and every run saves its prompt as `prompt.txt`.

## Grading

`tasks.py` grades each output file. It reads the OOXML parts directly with regular expressions and does not use a third-party library, so a grade does not depend on any one document library's interpretation of the file. Rerun `run.py --grade-only` to re-grade existing results.

### Content checks

Each task has its own content checks. For example: the required clause text is present, the table row sits directly above the total, or the PDF has 4–6 pages.

Tracked-change tasks also check three things:

- The output has tracked changes.
- Every revision (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`) is by the requested author.
- Rejecting every change gives back exactly the original text.

### Fidelity checks

Fidelity checks apply to tasks that edit an existing document, and they are deliberately strict. They capture whether a tool changes only what was asked or re-serializes the whole file.

- **Other parts byte-identical.** Every package part outside the set an edit may legitimately touch must still exist and be byte-for-byte identical to the input. That set is `word/document.xml`, the comments parts, `people.xml`, `[Content_Types].xml`, `document.xml.rels` and `settings.xml`. Parts that must not change include styles, numbering, headers, footers, theme, fonts and document properties.
  - A tool that loads `styles.xml` and writes it back with different attribute order, whitespace or namespace declarations fails this check, even if Word would render the file the same way.
- **At most N paragraphs changed.** The grader counts the input's `<w:p>` elements that no longer appear verbatim in the output. N is the number of paragraphs the task needs to change. For example, it is 4 for `redline3`, 3 for `long-doc-fix`, 1 for `heading-fix`, and for `bulk-rename` the number of paragraphs that mention "Provider".
  - Before comparing, the grader treats an empty element written as `<x></x>` and as `<x/>` as the same.
  - Any other byte difference in a paragraph counts as a change. That includes new `rsid` attributes, reordered attributes, split or merged runs, and changed properties.

A failed fidelity check does not necessarily mean the document looks wrong in Word. It means the tool changed bytes the task did not ask it to change. That matters for documents under review, for diffs, and for files that pass through other tools.

`comment-review` only checks that the other parts are byte-identical and that the text is unchanged. `compare`, `resolve-tracked`, `to-pdf` and `generate-memo` check content only, because they produce a new document.

## Isolation

The harness takes these steps to keep runs isolated and comparable:

- **Fresh directories.** Every run uses a fresh work directory that holds only that task's fixtures. In the `revise` arm, Claude Code and Codex get the skill inside the work directory (`.claude/skills`, `.agents/skills`), so no skill is installed globally.
- **No GUI.** `open` and `osascript` are stubbed out on `PATH`, and `soffice`/`libreoffice` always run with `--headless`. No agent can drive Word or any other app on the screen.
- **Sequential runs within an arm.** Baseline agents often write to fixed `/tmp` paths. The two arms run in parallel, except for Hermes with a shared OAuth login (see Credentials).
- **Claude Code.** User hooks are disabled. Allowed tools are `Bash, Read, Write, Edit, Glob, Grep, Skill`, with a cap of 80 turns. There is no web access.
- **Codex.** Runs with `--ignore-user-config` and `-c features.<name>=false` for hooks, apps, plugins, computer use, the in-app browser, browser use and image generation. Web search is disabled and notifications are off. Reasoning effort is medium. The shell, patches and images remain available.
- **Hermes.** Each arm gets its own profile, rendered from `hermes/config.yaml` into a temporary `HERMES_HOME` (never `~/.hermes`). In both profiles:
  - memory and user profiles are off, skill creation is off (`creation_nudge_interval: 0`), and there are no MCP servers;
  - telemetry and update checks are off;
  - Hermes's bundled skills are seeded as an interactive first run would;
  - toolsets are `terminal, file, code_execution, skills, todo, vision`, with no browser, web or computer use.

  The two profiles differ only in skills:
  - `revisebench` (revise arm) has `skills.disabled: [docx]`, and `skills.external_dirs` points at the installed revise skill.
  - `revisebase` (baseline arm) keeps every bundled skill, `docx` included.
- **Credentials.** Each agent gets exactly one credential and no other API key (see Credentials). Credential files are written owner-only, outside the results directory, and deleted when the run ends.
- **Redaction.** Before a record is written, any credential value is redacted from the agent's final message. Any file in the run directory that contains a credential is deleted, and the record lists it under `redacted_files`.

Local runs inherit your user-level Claude Code configuration, such as memory files or MCP servers. GitHub-hosted runners start from a clean home directory.

## Running locally

You need:

- Python 3.10+
- Node 20+
- the agent CLI you want to test
- for a fair baseline, the baseline toolset: `pip install python-docx lxml`, `pandoc`, and LibreOffice

```bash
cd bench
python3 run.py --agent claude --dry-run          # set up work dirs and profiles, print the commands
python3 run.py --agent claude --runs 4           # results in out/claude/
python3 run.py --agent codex --model gpt-6-sol
python3 run.py --agent hermes --provider xai-oauth --model <model> --hermes-auth-file ./hermes-auth.json
python3 summarize.py out/claude out/codex        # Markdown report
python3 run.py --grade-only --out out/claude     # re-grade without running agents
```

`--revise-spec` picks the `revise` package. `run.py` installs it into a temporary npm prefix with `npm install`. The default is `$REVISE_SPEC` if set, else `revise@latest`. It accepts:

- an npm spec: `revise@3.0.0`, `revise@next`;
- a tarball path or URL: `./revise-3.0.0.tgz`;
- several specs separated by commas, for an unpublished `revise` with unpublished dependencies: `./revise-3.0.0.tgz,./reviseio-engine-3.0.0.tgz`;
- `source`: build from this repository's checkout (see Building from source).

`--revise-bin PATH` uses an existing binary instead.

Other options: `--arms`, `--tasks`, `--runs` (default 4), `--timeout` (seconds per run, default 900), `--fixtures-dir`, `--out` (default `out/<agent>`).

### Results format

`run.py --out DIR` writes these files:

- **`DIR/meta.json`**: one per run of the harness. It holds:
  - the agent and its `--version`
  - model and provider
  - billing mode
  - `revise --version`
  - the npm spec, plus the resolved tarball URL and integrity of `revise` and `@reviseio/engine`
  - tasks, run count and timeout
  - start and finish times
  - the harness commit
  - host platform, Python, Node, LibreOffice, pandoc and python-docx versions
  - the GitHub Actions run URL
- **`DIR/<arm>/<task>/<n>/record.json`**: one per run, with checks, `info` (changed parts and paragraph counts), turns, wall time, cost, tokens, and the agent's final message (truncated to 1,500 characters).
- **`DIR/<arm>/<task>/<n>/transcript.jsonl`**: the agent's event stream (Claude Code and Codex), with credentials redacted. The record's `skill_read` says whether a revise-arm agent loaded the skill (Claude's Skill tool, a read of its `SKILL.md`, or `revise skill`), and `revise_calls` counts its `revise` commands. The summary reports how many revise-arm runs read the skill.
- **`DIR/<arm>/<task>/<n>/work/`**: the agent's output files.
- **`DIR/summary.json`**: all records.

## Running in GitHub Actions

`.github/workflows/benchmark.yml` runs only on manual dispatch (Actions → Benchmark → Run workflow). It has no `push`, `pull_request` or `pull_request_target` trigger, because the jobs hold credentials and the repository is public.

| Input | Default | Meaning |
|---|---|---|
| `agents` | `claude,codex,hermes` | One job per agent, in parallel |
| `models` | blank | `agent=model` pairs, e.g. `claude=opus,codex=gpt-6-sol`. If blank, the defaults are `sonnet`, `gpt-6-sol` and `gpt-5.6-terra`. |
| `tasks` | blank | Comma-separated task names. If blank, every task whose fixtures exist. |
| `runs` | `4` | Runs per task per arm |
| `revise_spec` | `revise@latest` | As `--revise-spec`. Tarballs must be URLs here. |
| `arms` | `revise,baseline` | Which arms to run |
| `skill_ref` | blank | With `revise_spec` `source`: take only `src/cli/skills/` from this branch or commit, for an A/B of skill text |
| `source_ref` | blank | With `revise_spec` `source`: build revise from `src/` at this branch or commit, to measure unreleased changes. The harness and graders stay on the workflow's ref. |
| `hermes_provider` | `openai-codex` | Hermes `--provider` |
| `publish` | off | Commit the aggregated results and open a pull request |

### Jobs

1. **Plan** validates the inputs.
2. **Bench** runs once per agent. It uses the `benchmark` environment, so the repository owner can require approval before any job that holds secrets starts.
   - It installs:
     - Node 22 and Python 3.12;
     - LibreOffice (`libreoffice-writer-nogui`), pandoc, python-docx and fonts;
     - the agent CLI. Versions are set by `CLAUDE_CODE_VERSION`, `CODEX_VERSION` and `HERMES_COMMIT` at the top of the workflow.
   - It then runs `run.py` and adds the Markdown summary to the job summary.
   - It uploads `bench/out/<agent>` as the `results-<agent>` artifact, which includes work directories and output documents.
3. **Aggregate** combines all agents into one summary and uploads `benchmark-summary`. That artifact holds `summary.md` plus `meta.json` and `records.json` for each agent, without the work directories.
4. **Publish** runs only if `publish` is set. It commits `benchmark-summary` under `bench/results/<date>-<sha>/` on a new branch and opens a pull request.
   - This is the only job with `contents: write` and `pull-requests: write`. Every other job has `contents: read`.
   - It needs the repository setting "Allow GitHub Actions to create and approve pull requests".

All third-party actions are pinned to full commit SHAs. Each job and long step has a timeout.

## Credentials

Add the secrets to the `benchmark` environment (Settings → Environments → benchmark), not as repository-wide secrets. Each bench job sees only its own agent's secrets.

| Secret | Agent | Billing | How to get it |
|---|---|---|---|
| `CLAUDE_CODE_OAUTH_TOKEN` | Claude Code | **Subscription** (Claude Pro/Max) | Run `claude setup-token` locally and paste the token. |
| `ANTHROPIC_API_KEY` | Claude Code | API (per token) | Anthropic Console |
| `CODEX_AUTH_JSON` | Codex | **Subscription** (ChatGPT plan) | The contents of `~/.codex/auth.json` after `codex login` with ChatGPT. The workflow writes it to `$CODEX_HOME/auth.json`. |
| `OPENAI_API_KEY` | Codex, Hermes (`openai-api`) | API (per token) | OpenAI platform. Codex receives it as `CODEX_API_KEY` for `codex exec`. |
| `HERMES_AUTH_JSON` | Hermes | **Subscription** for the OAuth providers `openai-codex` (ChatGPT) and `xai-oauth` (SuperGrok) | The contents of a Hermes `auth.json` (`$HERMES_HOME/auth.json`) after `hermes login` for that provider. |
| `XAI_API_KEY` | Hermes (`xai`) | API (per token) | xAI console |
| `OPENROUTER_API_KEY` | Hermes (`openrouter`) | API (per token) | OpenRouter |

`run.py` passes each agent a single credential, picked in this order:

- **Claude Code**:
  1. `CLAUDE_CODE_OAUTH_TOKEN`
  2. `ANTHROPIC_API_KEY`
  3. your local login (only when neither is set; local runs only)
- **Codex**:
  1. `$CODEX_HOME/auth.json` (default `~/.codex/auth.json`). It counts as a subscription when it holds ChatGPT tokens.
  2. `OPENAI_API_KEY`
- **Hermes**:
  1. the `auth.json` given by `--hermes-auth-file`, or `HERMES_AUTH_FILE` in the workflow
  2. the API key for the chosen provider (`OPENAI_API_KEY` for `openai-api`, `XAI_API_KEY` for `xai`, `OPENROUTER_API_KEY` for `openrouter`)

  OAuth providers require the auth file.

When a subscription credential is present, every API-key variable is removed from the agent's environment, so a run cannot silently fall back to API billing. Each run's `meta.json` records the billing mode it used, for example `subscription (CLAUDE_CODE_OAUTH_TOKEN)` or `api (OPENAI_API_KEY)`.

### OAuth refresh tokens

Codex and Hermes OAuth logins use single-use refresh tokens with reuse detection.

- **Refreshes make the stored secret stale.** When a CI run refreshes the login, the stored secret goes stale. If two copies of the same login refresh independently (CI and your laptop), the provider can revoke the login.
- **Use a separate login for CI.** Use a login dedicated to CI rather than the one on your workstation, and expect to refresh the secret from time to time.
- **Hermes shares one copy.** `run.py` copies the Hermes `auth.json` into both arm profiles. It runs the two arms alternately rather than in parallel, and copies refreshed tokens to the other profile and back to the `--hermes-auth-file` after every run.

## Cost and time

For reference, from runs of the full suite (10 tasks × 2 arms × 4 runs = 80 agent runs):

- **Claude Code** with Sonnet: about $19 at API prices. Median $0.21 (revise) and $0.26 (baseline) per run. About 25 minutes of wall time for the slower arm.
- **Codex** with gpt-6-sol: about $8 at list prices (computed from tokens). About 50 minutes for the slower arm. One run took 13 minutes.

On a subscription the dollar figures do not apply. Usage counts against the plan's limits instead.

Each run is capped at `--timeout` (15 minutes by default), and each bench job at 340 minutes. Hermes with an OAuth login runs its arms one after the other, which roughly doubles its wall time.

## Fixtures

| File | Provenance |
|---|---|
| `contract.docx` | A fictional services agreement, generated by `make_contract.py` with python-docx. With python-docx 1.2.0 the output is byte-identical. |
| `contract-v2.docx` | `contract.docx` after a negotiation pass by the Revise agent, exported with every change accepted. The agent's four comments remain. |
| `contract-tracked.docx` | The same pass as tracked changes. Accepting all of them gives `contract-v2.docx`'s text, and rejecting all of them gives `contract.docx`'s text. |
| `handbook.docx` | A fictional ~115-page employee handbook, written as raw OOXML by `make_handbook.py`, so it carries no bias toward any exporter. Deterministic: the script reproduces it byte for byte. |
| `nda.docx` | **Not included.** The NDA used by `comment-review` in earlier runs is a third-party test document. The task is skipped unless you supply an NDA as `fixtures/nda.docx` or through `--fixtures-dir`. |

The names and companies in all fixtures are fictional.

## Building from source

`--revise-spec source` (or the workflow input `revise_spec: source`) builds `revise` from the repository's own checkout instead of npm: `hooks/build-revise-from-source` runs `npm install`, `npm run build` and `npm pack ./dist`, exactly as the package is published. `@reviseio/engine` still comes from npm. The hook contract:

- **Call:** `build-revise-from-source <repo root> <output dir>`
- **Output:** print one npm tarball path per line: `revise`, plus any dependency not yet on npm.
- **Install:** `run.py` installs the tarballs together, as it does for a comma-separated spec.

In the workflow, any toolchain the build needs must be installed before the "Run the benchmark" step.

## Reproducing a published result

1. Open the result's `meta.json`. It names the harness commit, the agent version, the model, the `revise` package with its tarball URL and integrity, and the toolchain versions.
2. Check out the harness commit.
3. Install the same agent version.
4. Run `run.py` with the same `--model`, `--revise-spec`, `--runs` and `--tasks`.

You can re-grade any result directory with `run.py --grade-only --out DIR`, without running an agent. Agent output varies between runs, so compare pass counts and medians across several runs rather than single runs.

## License

MIT, as the rest of this repository. The fixtures are covered by the same license.
