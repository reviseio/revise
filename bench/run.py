#!/usr/bin/env python3
"""Benchmark: coding agents doing Word-document work with vs. without the
revise CLI.

Each task runs a headless coding agent in a fresh directory holding copies of
its fixtures, in two arms:

  revise    the revise skill installed for the agent, and `revise` on PATH
  baseline  neither: whatever the machine has (python-docx, pandoc,
            LibreOffice, shell tools, the agent's own bundled skills)

Every result is graded by script (tasks.py): the task's requirements, plus
fidelity checks (paragraphs and package parts the task did not ask to change
must come back byte-identical). Runs within an arm are sequential: baseline
agents often write to fixed /tmp paths, so concurrent runs would collide.

  python3 run.py --agent claude --runs 4 --out out/claude
  python3 run.py --agent codex --revise-spec ./revise-3.0.0.tgz,./reviseio-engine-3.0.0.tgz
  python3 run.py --agent hermes --provider xai-oauth --model <model> --hermes-auth-file auth.json
  python3 run.py --agent claude --dry-run       # set up, print commands, run no agent
  python3 run.py --grade-only --out out/claude  # re-grade existing runs

Hermes (Nous Research's Hermes Agent) loads skills from its profile, not from
the working directory, so each arm gets its own profile, rendered from
hermes/config.yaml into a temporary HERMES_HOME: `revisebench` sees the revise
skill (skills.external_dirs) with Hermes's bundled `docx` skill disabled, and
`revisebase` keeps the bundled skills. Both have memory, user profiles, and
skill creation turned off, so runs can't learn from each other.

See README.md for credentials and billing.
"""
import argparse, datetime, json, os, platform, re, shlex, shutil, subprocess
import sys, tempfile, threading, time
from pathlib import Path
from string import Template

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))  # run with `python3 -I` too: tasks.py sits next to this file

import tasks as task_defs  # noqa: E402
from tasks import FIXTURES, TASKS  # noqa: E402

AGENTS = ("claude", "codex", "hermes")
DEFAULT_MODELS = {"claude": "sonnet", "codex": "gpt-6-sol", "hermes": "gpt-5.6-terra"}
HERMES_DEFAULT_PROVIDER = "openai-codex"
HERMES_CODEX_BASE_URL = "https://chatgpt.com/backend-api/codex"
# Hermes providers that bill a subscription through an OAuth login in auth.json.
HERMES_OAUTH_PROVIDERS = {"openai-codex", "xai-oauth", "nous"}
# Hermes providers that bill an API key, and the variable each reads.
HERMES_PROVIDER_KEYS = {"openai-api": "OPENAI_API_KEY", "xai": "XAI_API_KEY", "openrouter": "OPENROUTER_API_KEY"}
HERMES_PROFILES = {"revise": "revisebench", "baseline": "revisebase"}

# Every credential an agent could bill against. None reaches an agent unless
# resolve_auth() passes it on deliberately.
CREDENTIAL_VARS = (
    "ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN", "CLAUDE_CODE_OAUTH_TOKEN",
    "OPENAI_API_KEY", "CODEX_API_KEY", "CODEX_AUTH_JSON",
    "XAI_API_KEY", "OPENROUTER_API_KEY", "HERMES_AUTH_JSON",
    "GITHUB_TOKEN", "GH_TOKEN", "NPM_TOKEN", "NODE_AUTH_TOKEN",
)

# The revise package spec `source` means "build revise from this repository's
# own source checkout": hooks/build-revise-from-source builds and packs it
# (see README.md, "Building from source").
SOURCE_BUILD_HOOK = HERE / "hooks" / "build-revise-from-source"


# ── setup ────────────────────────────────────────────────────────────────────

def first_line(command, env=None):
    try:
        proc = subprocess.run(command, capture_output=True, text=True, timeout=60, env=env,
                              stdin=subprocess.DEVNULL)
        return (proc.stdout or proc.stderr).strip().splitlines()[0] if (proc.stdout or proc.stderr).strip() else None
    except (OSError, subprocess.TimeoutExpired):
        return None


def build_revise_from_source(dest):
    """HOOK: build revise from this repository's source and return the paths of
    the npm tarballs to install (revise and any unpublished dependencies).

    The script is hooks/build-revise-from-source, called as
    `build-revise-from-source <repo root> <output dir>` and must print one
    tarball path per line on stdout."""
    if not SOURCE_BUILD_HOOK.exists():
        sys.exit(f"--revise-spec source: {SOURCE_BUILD_HOOK} is missing")
    # Through sh: a checkout or copy may not keep the executable bit.
    proc = subprocess.run(["sh", str(SOURCE_BUILD_HOOK), str(HERE.parent), str(dest)], capture_output=True, text=True)
    if proc.returncode != 0:
        sys.exit(f"--revise-spec source: the build hook failed (exit {proc.returncode}):\n{proc.stderr.strip()}")
    tarballs = [line.strip() for line in proc.stdout.splitlines() if line.strip()]
    if not tarballs:
        sys.exit("--revise-spec source: the build hook printed no tarballs")
    return tarballs


def install_revise(spec, tools_dir):
    """Install revise from npm into tools_dir and return (binary, provenance).

    spec is an npm spec (revise@latest, revise@3.0.0), a tarball path or URL,
    several of those separated by commas or spaces (for an unpublished revise
    with unpublished dependencies), or `source`."""
    if spec.strip() == "source":
        specs = build_revise_from_source(tools_dir / "source-build")
    else:
        specs = [s for s in re.split(r"[\s,]+", spec) if s]
    # Local tarballs resolve against the caller's directory.
    specs = [str(Path(s).resolve()) if not re.match(r"^[a-z]+:", s) and Path(s).exists() else s for s in specs]
    prefix = tools_dir / "npm"
    prefix.mkdir(parents=True, exist_ok=True)
    (prefix / "package.json").write_text('{"name": "revise-bench-tools", "private": true}\n')
    proc = subprocess.run(["npm", "install", "--prefix", str(prefix), "--no-audit", "--no-fund",
                           "--loglevel=error", *specs], capture_output=True, text=True)
    if proc.returncode != 0:
        sys.exit(f"npm install {' '.join(specs)} failed:\n{proc.stderr.strip()}")
    binary = prefix / "node_modules" / ".bin" / "revise"
    if not binary.exists():
        sys.exit(f"{' '.join(specs)} installed no `revise` binary")
    provenance = {"spec": spec, "npm_specs": specs}
    try:
        lock = json.loads((prefix / "package-lock.json").read_text())
        for name in ("revise", "@reviseio/engine"):
            entry = lock.get("packages", {}).get(f"node_modules/{name}")
            if entry:
                provenance[name] = {k: entry.get(k) for k in ("version", "resolved", "integrity")}
    except (OSError, json.JSONDecodeError):
        pass
    return binary, provenance


def hermes_install_dir(args):
    if args.hermes_install_dir:
        return Path(args.hermes_install_dir)
    out = subprocess.run(["hermes", "--version"], capture_output=True, text=True).stdout if shutil.which("hermes") else ""
    m = re.search(r"Install directory:\s*(\S+)", out)
    return Path(m.group(1)) if m else None


def setup_hermes(args, tmp, revise_bin):
    """Render the two arm profiles into fresh HERMES_HOMEs under tmp, install
    the revise skill for the revise arm, and seed Hermes's bundled skills (a
    one-shot `hermes -z` run never seeds them itself)."""
    skills_dir = tmp / "hermes-skills"
    if revise_bin:
        subprocess.run([str(revise_bin), "skill", "install", "--dir", str(skills_dir)], check=True, capture_output=True)
    model_block = f"  default: {args.model}\n  provider: {args.provider}"
    if args.provider == "openai-codex":
        model_block += f"\n  base_url: {HERMES_CODEX_BASE_URL}"
    install = hermes_install_dir(args)
    bundled = install / "skills" if install else None
    homes = {}
    for arm, profile in HERMES_PROFILES.items():
        home = tmp / "hermes" / profile
        home.mkdir(parents=True)
        os.chmod(home, 0o700)
        if arm == "revise":
            skills_block = f"  disabled:\n    - docx\n  external_dirs:\n    - {skills_dir}"
        else:
            skills_block = "  disabled: []"
        template = Template((HERE / "hermes/config.yaml").read_text())
        (home / "config.yaml").write_text(template.substitute(model_block=model_block, skills_block=skills_block))
        if args.hermes_auth_file and not args.dry_run:
            shutil.copy(args.hermes_auth_file, home / "auth.json")
            os.chmod(home / "auth.json", 0o600)
        if bundled and bundled.is_dir():
            # What Hermes's interactive first run does (tools/skills_sync.py).
            shutil.copytree(bundled, home / "skills", ignore=shutil.ignore_patterns("__pycache__"))
        else:
            print(f"warning: Hermes install directory not found; bundled skills not seeded into {home}", file=sys.stderr)
        homes[arm] = home
    return homes


def sync_hermes_auth(homes, source_arm, args):
    """OAuth refresh tokens are single-use: when one profile refreshes, every
    other copy (and the caller's file) must get the new tokens, or the next
    refresh reuses a spent token and the provider revokes the login."""
    src = homes[source_arm] / "auth.json"
    if not src.exists():
        return
    for arm, home in homes.items():
        if arm != source_arm and (home / "auth.json").read_bytes() != src.read_bytes():
            shutil.copy(src, home / "auth.json")
    if args.hermes_auth_file and Path(args.hermes_auth_file).read_bytes() != src.read_bytes():
        target = Path(args.hermes_auth_file)
        tmp_target = target.with_name(target.name + ".tmp")
        shutil.copy(src, tmp_target)
        os.chmod(tmp_target, 0o600)
        os.replace(tmp_target, target)


# ── credentials ──────────────────────────────────────────────────────────────

def scrubbed_env():
    return {k: v for k, v in os.environ.items() if k not in CREDENTIAL_VARS and not k.startswith("ACTIONS_")}


def resolve_auth(args):
    """The environment the agent runs with, and how it bills. Only the one
    credential the agent should use passes through; a subscription credential
    always wins, and when one is present every API key is withheld, so a run
    never falls back to API billing silently."""
    src, env = os.environ, scrubbed_env()
    if args.agent == "claude":
        if src.get("CLAUDE_CODE_OAUTH_TOKEN"):
            env["CLAUDE_CODE_OAUTH_TOKEN"] = src["CLAUDE_CODE_OAUTH_TOKEN"]
            return env, "subscription (CLAUDE_CODE_OAUTH_TOKEN)"
        if src.get("ANTHROPIC_API_KEY"):
            env["ANTHROPIC_API_KEY"] = src["ANTHROPIC_API_KEY"]
            return env, "api (ANTHROPIC_API_KEY)"
        return env, "local login"
    if args.agent == "codex":
        auth = Path(src.get("CODEX_HOME") or Path.home() / ".codex") / "auth.json"
        if auth.exists():
            try:
                tokens = json.loads(auth.read_text()).get("tokens")
            except (OSError, json.JSONDecodeError, AttributeError):
                tokens = None
            return env, "subscription (CODEX_HOME/auth.json)" if tokens else "api (CODEX_HOME/auth.json)"
        key = src.get("CODEX_API_KEY") or src.get("OPENAI_API_KEY")
        if key:
            env["CODEX_API_KEY"] = key  # `codex exec` reads CODEX_API_KEY
            return env, "api (OPENAI_API_KEY)"
        return env, None
    if args.hermes_auth_file:
        kind = "subscription" if args.provider in HERMES_OAUTH_PROVIDERS else "auth file"
        return env, f"{kind} (auth.json, provider {args.provider})"
    var = HERMES_PROVIDER_KEYS.get(args.provider)
    if var and src.get(var):
        env[var] = src[var]
        return env, f"api ({var})"
    return env, None


def secret_values(args):
    """Strings that must never appear in a published record or work file."""
    values = {v for k, v in os.environ.items() if k in CREDENTIAL_VARS and v}

    def walk(node, key=""):
        if isinstance(node, dict):
            for k, v in node.items():
                walk(v, k)
        elif isinstance(node, list):
            for v in node:
                walk(v, key)
        elif isinstance(node, str) and re.search(r"token|key|secret|password", key, re.I):
            values.add(node)

    files = [args.hermes_auth_file, Path(os.environ.get("CODEX_HOME") or Path.home() / ".codex") / "auth.json"]
    for path in files:
        try:
            walk(json.loads(Path(path).read_text()))
        except (OSError, TypeError, ValueError):
            pass
    return sorted((v for v in values if len(v) >= 12), key=len, reverse=True)


def redact(text, secrets):
    for value in secrets:
        text = text.replace(value, "[REDACTED]")
    return text


def scrub_work_dir(work, secrets):
    """Delete any file an agent wrote that contains a credential (e.g. an
    environment dump), so results can be uploaded and published."""
    removed = []
    needles = [s.encode() for s in secrets]
    for path in work.rglob("*"):
        try:
            if path.is_file() and not path.is_symlink() and path.stat().st_size < 50_000_000:
                data = path.read_bytes()
                if any(n in data for n in needles):
                    path.unlink()
                    removed.append(str(path.relative_to(work)))
        except OSError:
            pass
    return removed


# ── runners ──────────────────────────────────────────────────────────────────

# Standard gpt-6-sol rates (USD per million tokens: input, cached input, output).
CODEX_PRICES = {"gpt-6-sol": (2.00, 0.20, 10.00)}
CODEX_TOOL_ITEMS = {"command_execution", "file_change", "mcp_tool_call", "web_search"}


def claude_command(prompt, args, run_dir):
    return ["claude", "-p", prompt, "--model", args.model, "--output-format", "json",
            "--settings", '{"disableAllHooks": true}',  # no user hooks (sounds)
            "--max-turns", "80", "--allowedTools", "Bash,Read,Write,Edit,Glob,Grep,Skill"]


def codex_command(prompt, args, work):
    # Stock Codex: no user config (its MCP servers, plugins, hooks, and
    # notification sounds), no ChatGPT account apps, no web, browser, or
    # computer use. Shell, patches, and images remain.
    return ["codex", "exec", "-m", args.model, "--json", "--skip-git-repo-check",
            "--ignore-user-config", "-c", 'model_reasoning_effort="medium"',
            *[arg for feature in ("hooks", "apps", "plugins", "computer_use",
                                  "in_app_browser", "browser_use", "image_generation")
              for arg in ("-c", f"features.{feature}=false")],
            "-c", 'web_search="disabled"', "-c", "notify=[]",
            "--dangerously-bypass-approvals-and-sandbox", "-C", str(work), prompt]


def hermes_command(prompt, args, run_dir):
    # The tools a coding agent has (as the Claude Code arm's --allowedTools):
    # no computer use, browser, or web, which could drive apps on the screen.
    return ["hermes", "-z", prompt, "--usage-file", str(run_dir / "usage.json"),
            "-t", "terminal,file,code_execution,skills,todo,vision",
            "-m", args.model, "--provider", args.provider]


def run_claude(command, work, env, args):
    try:
        proc = subprocess.run(command, cwd=work, env=env, capture_output=True, text=True,
                              timeout=args.timeout, stdin=subprocess.DEVNULL)
        raw = proc.stdout
    except subprocess.TimeoutExpired:
        raw = ""
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"is_error": True, "result": raw[-2000:]}


def run_codex(command, work, env, args):
    """Headless Codex: turns = tool steps + the final answer; cost from the
    token usage at the model's standard rates."""
    try:
        proc = subprocess.run(command, cwd=work, env=env, capture_output=True, text=True,
                              timeout=args.timeout, stdin=subprocess.DEVNULL)
        lines = proc.stdout.splitlines()
    except subprocess.TimeoutExpired as error:
        out = error.stdout or ""
        lines = (out.decode(errors="replace") if isinstance(out, bytes) else out).splitlines()
    steps, final, usage = 0, "", {"input": 0, "cached": 0, "output": 0}
    for line in lines:
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        item = event.get("item") or {}
        if event.get("type") == "item.completed" and item.get("type") in CODEX_TOOL_ITEMS:
            steps += 1
        if event.get("type") == "item.completed" and item.get("type") == "agent_message":
            final = item.get("text", "")
        if event.get("type") == "turn.completed":
            u = event.get("usage") or {}
            usage["input"] += u.get("input_tokens", 0)
            usage["cached"] += u.get("cached_input_tokens", 0)
            usage["output"] += u.get("output_tokens", 0)
    usage["total_tokens"] = usage["input"] + usage["output"]
    price_in, price_cached, price_out = CODEX_PRICES.get(args.model, (0, 0, 0))
    cost = ((usage["input"] - usage["cached"]) * price_in + usage["cached"] * price_cached
            + usage["output"] * price_out) / 1_000_000
    return {"num_turns": steps + 1, "total_cost_usd": cost if args.model in CODEX_PRICES else None,
            "is_error": not final, "session_id": None, "result": final, "usage": usage}


def run_hermes(command, work, env, args, run_dir):
    """Headless Hermes (`hermes -z`): turns = API calls; cost from its usage
    report when the provider prices it, else None (tokens are recorded)."""
    try:
        proc = subprocess.run(command, cwd=work, env=env, capture_output=True, text=True,
                              timeout=args.timeout, stdin=subprocess.DEVNULL)
        final = proc.stdout
    except subprocess.TimeoutExpired:
        final = ""
    try:
        usage = json.loads((run_dir / "usage.json").read_text())
    except (OSError, json.JSONDecodeError):
        usage = {}
    cost = usage.get("estimated_cost_usd") if usage.get("cost_status") not in ("included", None) else None
    return {"num_turns": usage.get("api_calls"), "total_cost_usd": cost,
            "is_error": bool(usage.get("failed")) or not final.strip(),
            "session_id": usage.get("session_id"), "result": final,
            "usage": {k: usage.get(k) for k in ("input_tokens", "cache_read_tokens",
                                                "output_tokens", "reasoning_tokens", "total_tokens")}}


def prepare_run(arm, task, n, ctx):
    """Fresh run directory: fixtures, no-GUI stubs, and (revise arm) the revise
    binary and skill. Returns (run_dir, work, env, command)."""
    args = ctx["args"]
    files, _, prompt = TASKS[task]
    run_dir = ctx["out_root"] / arm / task / str(n)
    work = run_dir / "work"
    shutil.rmtree(run_dir, ignore_errors=True)
    work.mkdir(parents=True)
    for name in files:
        shutil.copy(FIXTURES[name], work / name)
    (run_dir / "prompt.txt").write_text(prompt)
    env = dict(ctx["env"])
    # A benchmark machine is headless: no agent may drive Word or any other
    # GUI app (docx2pdf and `open` launch Word on a Mac, on the user's screen).
    no_gui = run_dir / "no-gui"
    no_gui.mkdir()
    for name in ("open", "osascript"):
        stub = no_gui / name
        stub.write_text("#!/bin/sh\necho \"$(basename \"$0\"): GUI applications are not available on this headless machine.\" >&2\nexit 1\n")
        stub.chmod(0o755)
    # LibreOffice stays available, but always headless: no window on screen.
    real_soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if real_soffice:
        for name in ("soffice", "libreoffice"):
            stub = no_gui / name
            stub.write_text(f'#!/bin/sh\nexec "{real_soffice}" --headless "$@"\n')
            stub.chmod(0o755)
    # The baseline arm has no revise: a global install sits next to node on
    # PATH, and one baseline found and used it.
    if arm == "baseline":
        stub = no_gui / "revise"
        stub.write_text("#!/bin/sh\necho \"revise: command not found\" >&2\nexit 127\n")
        stub.chmod(0o755)
    env["PATH"] = f"{no_gui}:{env['PATH']}"
    if arm == "revise":
        bin_dir = run_dir / "bin"
        bin_dir.mkdir()
        (bin_dir / "revise").symlink_to(Path(ctx["revise_bin"]).resolve())
        env["PATH"] = f"{bin_dir}:{env['PATH']}"
        if args.agent != "hermes":  # Hermes reads the skill from its profile (setup_hermes)
            skills = ".agents/skills" if args.agent == "codex" else ".claude/skills"
            subprocess.run([str(bin_dir / "revise"), "skill", "install", "--dir", skills], cwd=work, env=env,
                           check=True, capture_output=True)
    if args.agent == "hermes":
        env["HERMES_HOME"] = str(ctx["hermes_homes"][arm])
        command = hermes_command(prompt, args, run_dir)
    elif args.agent == "codex":
        command = codex_command(prompt, args, work)
    else:
        command = claude_command(prompt, args, run_dir)
    return run_dir, work, env, command


def run_one(arm, task, n, ctx):
    args, secrets = ctx["args"], ctx["secrets"]
    run_dir, work, env, command = prepare_run(arm, task, n, ctx)
    started = time.time()
    if args.agent == "hermes":
        result = run_hermes(command, work, env, args, run_dir)
        if ctx["hermes_auth_lock"]:
            sync_hermes_auth(ctx["hermes_homes"], arm, args)
    elif args.agent == "codex":
        result = run_codex(command, work, env, args)
    else:
        result = run_claude(command, work, env, args)
    wall = time.time() - started
    removed = scrub_work_dir(run_dir, secrets)
    try:
        checks, info = TASKS[task][1](work)
    except Exception as error:  # a grader crash is a failed run, not a harness failure
        checks, info = {"grader error": False}, {"error": repr(error)}
    record = {
        "agent": args.agent, "model": args.model, "provider": args.provider if args.agent == "hermes" else None,
        "arm": arm, "task": task, "run": n, "wall_s": round(wall, 1),
        "turns": result.get("num_turns"), "cost_usd": result.get("total_cost_usd"),
        "is_error": result.get("is_error"), "session_id": result.get("session_id"),
        "passed": all(checks.values()), "checks": checks, "info": info,
        "final_message": redact((result.get("result") or "")[:1500], secrets),
        "usage": result.get("usage"),
    }
    if removed:
        record["redacted_files"] = removed
    (run_dir / "record.json").write_text(json.dumps(record, indent=2))
    print(f"[{arm}] {task} #{n}: {'PASS' if record['passed'] else 'FAIL'} "
          f"{record['turns']} turns {record['wall_s']}s ${record['cost_usd'] or 0:.2f} "
          f"{[k for k, v in checks.items() if not v]}", flush=True)
    return record


# ── metadata ─────────────────────────────────────────────────────────────────

def collect_meta(args, revise_bin, revise_provenance, billing, env):
    agent_version = first_line([args.agent, "--version"], env=env) if shutil.which(args.agent) else None
    py_docx = first_line(["python3", "-c", "import importlib.metadata as m; print(m.version('python-docx'))"], env=env)
    head = subprocess.run(["git", "-C", str(HERE), "rev-parse", "--verify", "-q", "HEAD"], capture_output=True, text=True)
    harness = {"commit": head.stdout.strip() if head.returncode == 0 else None}
    if harness["commit"]:
        harness["dirty"] = bool(subprocess.run(["git", "-C", str(HERE), "status", "--porcelain", "--", "."],
                                               capture_output=True, text=True).stdout.strip())
    return {
        "agent": args.agent, "agent_version": agent_version,
        "model": args.model, "provider": args.provider if args.agent == "hermes" else None,
        "billing": billing,
        "revise_version": first_line([str(revise_bin), "--version"]) if revise_bin else None,
        "revise_package": revise_provenance,
        "arms": args.arms.split(","), "tasks": args.tasks.split(","), "runs": args.runs,
        "timeout_s": args.timeout,
        "started_at": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
        "harness": harness,
        "host": {
            "platform": platform.platform(), "python": platform.python_version(),
            "node": first_line(["node", "--version"]),
            "libreoffice": first_line(["soffice", "--version"]) if shutil.which("soffice") else None,
            "pandoc": first_line(["pandoc", "--version"]) if shutil.which("pandoc") else None,
            "python_docx": py_docx if py_docx and re.match(r"^\d", py_docx) else None,
            "github_run": os.environ.get("GITHUB_RUN_ID") and
            f"{os.environ.get('GITHUB_SERVER_URL')}/{os.environ.get('GITHUB_REPOSITORY')}/actions/runs/{os.environ['GITHUB_RUN_ID']}",
        },
    }


# ── main ─────────────────────────────────────────────────────────────────────

def print_table(records, tasks, arms):
    print("\ntask              arm       pass   turns   wall s   cost $   tokens")
    for task in tasks:
        for arm in arms:
            rs = [r for r in records if r["task"] == task and r["arm"] == arm]
            if not rs:
                continue
            avg = lambda k: sum((r[k] or 0) for r in rs) / len(rs)  # noqa: E731
            tokens = sum(((r.get("usage") or {}).get("total_tokens") or 0) for r in rs) / len(rs)
            print(f"{task:17} {arm:9} {sum(r['passed'] for r in rs)}/{len(rs)}   {avg('turns'):5.1f}   "
                  f"{avg('wall_s'):6.0f}   {avg('cost_usd'):5.2f}   {tokens:8.0f}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--agent", choices=AGENTS, default="claude")
    ap.add_argument("--model", default=None,
                    help="default: sonnet (claude), gpt-6-sol (codex), gpt-5.6-terra (hermes)")
    ap.add_argument("--provider", default=None,
                    help=f"hermes only, default {HERMES_DEFAULT_PROVIDER}: e.g. xai-oauth, openai-api, xai, openrouter")
    ap.add_argument("--revise-spec", default=os.environ.get("REVISE_SPEC", "revise@latest"),
                    help="npm spec, tarball path or URL, several of those (comma-separated), or `source` "
                         "(default: $REVISE_SPEC or revise@latest)")
    ap.add_argument("--revise-bin", default=None, help="use this revise binary instead of installing --revise-spec")
    ap.add_argument("--hermes-auth-file", default=os.environ.get("HERMES_AUTH_FILE"),
                    help="Hermes auth.json to copy into both profiles; refreshed tokens are written back to it")
    ap.add_argument("--hermes-install-dir", default=None, help="Hermes source checkout (to seed bundled skills)")
    ap.add_argument("--arms", default="revise,baseline")
    ap.add_argument("--tasks", default=None, help="comma-separated (default: every task whose fixtures exist)")
    ap.add_argument("--runs", type=int, default=4)
    ap.add_argument("--timeout", type=int, default=900, help="seconds per agent run")
    ap.add_argument("--fixtures-dir", default=str(task_defs.FIXTURES_DIR))
    ap.add_argument("--out", default=None, help="default: bench/out/<agent>")
    ap.add_argument("--dry-run", action="store_true", help="set up work dirs and profiles, print the agent commands")
    ap.add_argument("--grade-only", action="store_true", help="re-grade existing runs in --out")
    args = ap.parse_args()

    task_defs.set_fixtures_dir(args.fixtures_dir)
    if args.agent == "hermes":
        args.provider = args.provider or HERMES_DEFAULT_PROVIDER
    args.model = args.model or DEFAULT_MODELS[args.agent]
    out_root = Path(args.out or HERE / "out" / args.agent).resolve()
    available = [t for t in TASKS if all(FIXTURES[f].exists() for f in TASKS[t][0])]
    if args.tasks:
        unknown = [t for t in args.tasks.split(",") if t not in TASKS]
        missing = [t for t in args.tasks.split(",") if t in TASKS and t not in available]
        if unknown or missing:
            sys.exit(f"unknown tasks {unknown}; tasks missing fixtures {missing}")
    else:
        for t in TASKS:
            if t not in available:
                print(f"note: skipping {t}: fixture(s) {[f for f in TASKS[t][0] if not FIXTURES[f].exists()]} "
                      f"not in {args.fixtures_dir}", file=sys.stderr)
        args.tasks = ",".join(available)
    arms, tasks = args.arms.split(","), args.tasks.split(",")

    if args.grade_only:
        records = []
        for path in sorted(out_root.glob("*/*/*/record.json")):
            record = json.loads(path.read_text())
            record["checks"], record["info"] = TASKS[record["task"]][1](path.parent / "work")
            record["passed"] = all(record["checks"].values())
            path.write_text(json.dumps(record, indent=2))
            records.append(record)
        (out_root / "summary.json").write_text(json.dumps(records, indent=2))
        print_table(records, list(dict.fromkeys(r["task"] for r in records)), arms)
        return

    env, billing = resolve_auth(args)
    if billing is None and not args.dry_run:
        sys.exit(f"no credentials for {args.agent}; see README.md (Credentials)")
    if args.agent == "hermes" and args.provider in HERMES_OAUTH_PROVIDERS and not args.hermes_auth_file and not args.dry_run:
        sys.exit(f"provider {args.provider} logs in through auth.json: pass --hermes-auth-file")
    if not args.dry_run and not shutil.which(args.agent):
        sys.exit(f"`{args.agent}` is not on PATH")

    out_root.mkdir(parents=True, exist_ok=True)
    # Credentials and installed tools live outside the results directory, which
    # is uploaded as-is, and are removed when the run ends.
    tmp = Path(tempfile.mkdtemp(prefix="revise-bench-", dir=os.environ.get("RUNNER_TEMP")))
    os.chmod(tmp, 0o700)
    try:
        if args.revise_bin:
            revise_bin, provenance = Path(args.revise_bin).resolve(), {"spec": None, "binary": str(args.revise_bin)}
        elif "revise" in arms or args.dry_run:
            revise_bin, provenance = install_revise(args.revise_spec, tmp / "tools")
        else:
            revise_bin, provenance = None, None
        ctx = {"args": args, "out_root": out_root, "env": env, "revise_bin": revise_bin,
               "secrets": secret_values(args), "hermes_homes": None,
               "hermes_auth_lock": args.agent == "hermes" and bool(args.hermes_auth_file)}
        if args.agent == "hermes":
            ctx["hermes_homes"] = setup_hermes(args, tmp, revise_bin)
        meta = collect_meta(args, revise_bin, provenance, billing, env)
        (out_root / "meta.json").write_text(json.dumps(meta, indent=2))
        print(f"{args.agent} {meta['agent_version'] or '(not installed)'} · model {args.model}"
              f"{' · provider ' + args.provider if args.agent == 'hermes' else ''} · billing {billing or 'none'} · "
              f"revise {meta['revise_version']}", flush=True)

        if args.dry_run:
            for arm in arms:
                for task in tasks:
                    for n in range(1, args.runs + 1):
                        run_dir, work, env_run, command = prepare_run(arm, task, n, ctx)
                        shown = [f'"$(cat {run_dir / "prompt.txt"})"' if c == TASKS[task][2] else shlex.quote(c) for c in command]
                        extra = f"HERMES_HOME={shlex.quote(env_run['HERMES_HOME'])} " if "HERMES_HOME" in env_run else ""
                        prepended = env_run["PATH"][:len(env_run["PATH"]) - len(ctx["env"]["PATH"])]
                        print(f"[{arm}] {task} #{n}\n  cd {shlex.quote(str(work))} && PATH={shlex.quote(prepended)}\"$PATH\" "
                              f"{extra}{' '.join(shown)}")
            print(f"\ndry run: {len(arms) * len(tasks) * args.runs} runs prepared under {out_root}; no agent was started")
            return

        records, lock = [], threading.Lock()

        def run_and_collect(arm, task, n):
            record = run_one(arm, task, n, ctx)
            with lock:
                records.append(record)

        if ctx["hermes_auth_lock"]:
            # One OAuth login shared by both profiles: run the arms alternately
            # so its single-use refresh token is never spent twice.
            for task in tasks:
                for n in range(1, args.runs + 1):
                    for arm in arms:
                        run_and_collect(arm, task, n)
        else:
            def arm_worker(arm):
                for task in tasks:
                    for n in range(1, args.runs + 1):
                        run_and_collect(arm, task, n)
            threads = [threading.Thread(target=arm_worker, args=(arm,)) for arm in arms]
            [t.start() for t in threads]
            [t.join() for t in threads]
        meta["finished_at"] = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")
        (out_root / "meta.json").write_text(json.dumps(meta, indent=2))
        (out_root / "summary.json").write_text(json.dumps(records, indent=2))
        print_table(records, tasks, arms)
    finally:
        if args.dry_run:
            print(f"profiles and tools kept for inspection in {tmp} (no credentials copied)")
        else:
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
