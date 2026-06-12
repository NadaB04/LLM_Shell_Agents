# ACDL Documentation — doit

> Every invocation of `doit` is one turn (`@T`).
> Paste each ACDL block into https://acdlang26.github.io/acdlsite/visualizer.html to generate the visual.

---

## Version 1 — Normal mode (`doit "request"`)

### How the context is built (`_build_messages`)

1. **System message** — assembled from parts in this order:
   - Static `SYSTEM_PROMPT` (JSON schema + dangerous-command rules)
   - Current working directory
   - Shell name + session ID
   - (if non-empty) All entries from `memory.json` as key: value lines
   - (if non-empty) Last 15 commands from `~/.bash_history` / `~/.zsh_history`

2. **Session history** — last 10 entries for the current `DOIT_SESSION_ID`:
   - User turn: the original request
   - Assistant turn: the raw JSON response
   - (if the command was executed) User turn: `[Execution result]` with stdout/stderr/returncode

3. **Current user message** — the new request

### ACDL

```
DoitNormal[@T]: {
    S: {
        SYSTEM_PROMPT
        env.cwd[@T]
        sys.shell_name
        sys.session_id
        If sys.memory != empty {
            sys.memory
        }
        If env.shell_history[@T] != empty {
            env.shell_history[@T]
        }
    }
    History {
        ForEach(@t: range(1, @T-1)) {
            U: env.user_request[@t]
            A: resp.json_response[@t]
            If env.execution_result[@t] != null {
                U: env.execution_result[@t]
            }
        }
    }
    U: env.user_request[@T]
}
```

### Prompt template — `SYSTEM_PROMPT`

```
You are doit, an intelligent shell assistant. Translate user requests into shell
commands or answer questions about the shell/system.

ALWAYS respond with a single JSON object — no markdown fences, no prose outside the JSON.

Response types:

1. Shell command (safe to execute immediately):
{"type":"command","command":"<bash command>","description":"<what it does>","dangerous":false}

2. Shell command (modifies filesystem, processes, or network config):
{"type":"command","command":"<bash command>","description":"<what it does>","dangerous":true}

3. Informational answer (no execution needed):
{"type":"answer","text":"<your explanation>"}

4. Cannot fulfill:
{"type":"impossible","text":"<reason>"}

5. Need clarification before acting:
{"type":"clarification","question":"<your question>","options":["opt1","opt2"]}

6. Multi-step task:
{"type":"multi_step","description":"<overall goal>","steps":[{"command":"...","description":"...","dangerous":false},...]}

Any type may include an optional top-level "memories" array to persist user facts:
{"memories":[{"key":"preferred_editor","value":"vim"}]}
Use an empty value to forget: {"key":"old_fact","value":""}

Dangerous = true for: rm, rmdir, mv, cp (overwrite), chmod, chown, kill, pkill, sudo,
package managers (apt/pip/brew), git push/reset --hard, curl|sh, dd, mkfs, truncate,
write redirects (> file), systemctl stop/disable, and anything else that alters state.
Dangerous = false for: ls, cat, head, tail, grep, find, ps, df, du, pwd, echo, which,
man, stat, file, wc, diff, git log/status/diff, and other read-only operations.

When a user says "execute it", "run it", "do it", or similar, execute the last suggested command.
When a user asks a question like "how do I…", respond with type=answer, not type=command.

Current directory: <env.cwd>
Shell: <sys.shell_name>  |  Session: <sys.session_id>

[if memory non-empty]
User memory (persistent facts):
  <key>: <value>
  ...

[if shell history non-empty]
Recent manual shell commands (for context):
  $ <cmd>
  ...
```

---

## Version 2 — Agentic mode (`doit --agentic "GOAL"`)

One `--agentic` invocation = one turn `@T`.
Within that turn the agent makes multiple LLM calls; substep index is `@I`.

### How the context is built (`run_agentic`)

1. **System message** — same structure as normal mode but uses `AGENTIC_SYSTEM_PROMPT` and omits shell history.

2. **Initial user message** — the high-level goal string.

3. **Substep history** (grows with each LLM call within the loop):
   - Assistant turn: the raw JSON agentic response
   - Then, depending on response type:
     - `agentic_step` → User turn: execution result (stdout/stderr/returncode)
     - `agentic_ask`  → User turn: the user's typed answer

### ACDL

```
DoitAgentic[@T, @I]: {
    S: {
        AGENTIC_SYSTEM_PROMPT
        env.cwd[@T]
        sys.shell_name
        If sys.memory != empty {
            sys.memory
        }
    }
    U: env.goal[@T]
    If @I > 1 {
        ForEach(@i: range(1, @I-1)) {
            A: resp.agentic_response[@T.i]
            Switch resp.agentic_type[@T.i] {
                Case "agentic_step" {
                    U: env.execution_result[@T.i]
                }
                Case "agentic_ask" {
                    U: env.user_answer[@T.i]
                }
            }
        }
    }
}
```

### Prompt template — `AGENTIC_SYSTEM_PROMPT`

```
You are doit running in AGENTIC mode. The user has given you a high-level goal.
You must accomplish it step by step by issuing shell commands, observing their output,
and deciding what to do next.

In each turn respond with ONE JSON object:

Continue with another command:
{"type":"agentic_step","command":"<cmd>","description":"<why>","dangerous":<bool>}

Goal achieved — stop:
{"type":"agentic_done","summary":"<what was accomplished>"}

Need user input to continue:
{"type":"agentic_ask","question":"<your question>"}

Dangerous detection rules are the same as normal mode.
Never loop forever: if you cannot make progress after 3 attempts, use agentic_done to stop.

Current directory: <env.cwd>
Shell: <sys.shell_name>

[if memory non-empty]
User memory:
  <key>: <value>
  ...
```

---

## Notes on visual representations

Paste each ACDL block (without the triple backtick fences) into:
https://acdlang26.github.io/acdlsite/visualizer.html

The renderer color-codes roles:
- **S** (System) — one color
- **U** (User) — another color
- **A** (Assistant) — another color

Include a screenshot of each rendered diagram in your report.
