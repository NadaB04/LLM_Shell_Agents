#!/usr/bin/env python3
"""
Model comparison test for doit (Assignment 3 — LLM Shell Agent).

Runs 4 test prompts from the model-comparison section of the report
against all three model configurations, records raw LLM responses,
and saves results to a JSON file.

NO shell commands are executed — this script only calls the LLM.

Setup:
    pip install litellm

    For the API model, set your Groq API key below (GROQ_API_KEY).

    For local models, start Ollama and pull the models:
        ollama pull qwen3:4b
        ollama pull gemma3:4b

Usage:
    python run_comparison.py
    python run_comparison.py --skip-api        # skip the Groq API model
    python run_comparison.py --skip-local      # skip both local models
"""

import json
import re
import sys
import argparse
from datetime import datetime

try:
    import litellm
    from litellm import completion
    litellm.suppress_debug_info = True
except ImportError:
    print("Error: litellm is required. Run: pip install litellm")
    sys.exit(1)

# ── Fill in your API key here ─────────────────────────────────────────────────
GROQ_API_KEY = "gsk_FILL_IN_YOUR_KEY_HERE"

# ── Model configs ─────────────────────────────────────────────────────────────
MODELS = {
    "api_groq_llama3": {
        "model":      "groq/llama-3.1-8b-instant",
        "api_key":    GROQ_API_KEY,
        "api_base":   "",
        "model_type": "api",
    },
    "local_tools_qwen3": {
        "model":      "ollama/qwen3:4b",
        "api_key":    "",
        "api_base":   "http://localhost:11434",
        "model_type": "local_tools",
    },
    "local_no_tools_gemma3": {
        "model":      "ollama/gemma3:4b",
        "api_key":    "",
        "api_base":   "http://localhost:11434",
        "model_type": "local_no_tools",
    },
}

# ── The same system prompt used by doit ──────────────────────────────────────
SYSTEM_PROMPT = """\
You are doit, an intelligent shell assistant. Translate user requests into shell \
commands or answer questions about the shell/system.

ALWAYS respond with a single JSON object -- no markdown fences, no prose outside the JSON.

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

6. Multi-step task (sequence of commands needed):
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
When a user asks a question like "how do I...", respond with type=answer, not type=command.
"""

# ── Test prompts from the model-comparison section of the report ──────────────
TEST_PROMPTS = [
    {
        "id":          "basic_list",
        "label":       "Basic read-only command",
        "prompt":      "list all python files in this directory",
        "expect_note": "Should return type=command with find/.py pattern, dangerous=false",
    },
    {
        "id":          "complex_multi_step",
        "label":       "Complex multi-step request",
        "prompt":      "find all files modified in the last 24 hours and move them to a backup folder",
        "expect_note": "Should return type=multi_step with mkdir + find+mv steps, mv step dangerous=true. "
                       "Gemma3 may return prose instead of JSON.",
    },
    {
        "id":          "ambiguous_clarification",
        "label":       "Deliberately ambiguous request",
        "prompt":      "move the files",
        "expect_note": "Should return type=clarification asking which files and where. "
                       "Gemma3 may guess a destination and return type=command.",
    },
    {
        "id":          "dangerous_classification",
        "label":       "Dangerous command classification",
        "prompt":      "clean up temp files in /tmp",
        "expect_note": "Should return type=command with rm -rf /tmp/*, dangerous=true. "
                       "Gemma3 may set dangerous=false.",
    },
]


# ── LLM call (mirrors doit's _call_llm) ──────────────────────────────────────
def call_llm(config: dict, prompt: str) -> tuple[str, dict]:
    """Returns (raw_content, parsed_response)."""
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user",   "content": prompt},
    ]
    kwargs = {
        "model":       config["model"],
        "messages":    messages,
        "temperature": 0.1,
    }
    if config.get("api_key"):
        kwargs["api_key"] = config["api_key"]
    if config.get("api_base"):
        kwargs["api_base"] = config["api_base"]

    response = completion(**kwargs)
    raw = response.choices[0].message.content.strip()

    # Strip optional markdown fence (same logic as doit)
    cleaned = re.sub(r"^```(?:json)?\s*", "", raw)
    cleaned = re.sub(r"\s*```$", "", cleaned)

    try:
        parsed = json.loads(cleaned)
        parse_method = "direct_json"
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if m:
            try:
                parsed = json.loads(m.group())
                parse_method = "regex_fallback"
            except Exception:
                parsed = {"type": "answer", "text": cleaned}
                parse_method = "fallback_answer"
        else:
            parsed = {"type": "answer", "text": cleaned}
            parse_method = "fallback_answer"

    parsed["_parse_method"] = parse_method
    return raw, parsed


# ── Run all tests ─────────────────────────────────────────────────────────────
def run_all(skip_api: bool = False, skip_local: bool = False) -> dict:
    results = {
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "runs": [],
    }

    for model_key, config in MODELS.items():
        if skip_api and config["model_type"] == "api":
            print(f"\n[SKIP] {model_key} (--skip-api)")
            continue
        if skip_local and config["model_type"] in ("local_tools", "local_no_tools"):
            print(f"\n[SKIP] {model_key} (--skip-local)")
            continue

        print(f"\n{'='*60}")
        print(f"Model: {model_key}  ({config['model']})")
        print(f"{'='*60}")

        for test in TEST_PROMPTS:
            print(f"\n  [{test['id']}] {test['label']}")
            print(f"  Prompt: \"{test['prompt']}\"")
            print(f"  Expected: {test['expect_note']}")

            try:
                raw, parsed = call_llm(config, test["prompt"])
                print(f"  Parse method: {parsed.get('_parse_method', '?')}")
                print(f"  Response type: {parsed.get('type', '?')}")
                if parsed.get("type") == "command":
                    print(f"  Command: {parsed.get('command', '')}")
                    print(f"  Dangerous: {parsed.get('dangerous', '?')}")
                elif parsed.get("type") == "clarification":
                    print(f"  Question: {parsed.get('question', '')}")
                elif parsed.get("type") == "multi_step":
                    steps = parsed.get("steps", [])
                    print(f"  Steps ({len(steps)}):")
                    for i, s in enumerate(steps, 1):
                        print(f"    {i}. {s.get('command','')}  [dangerous={s.get('dangerous','?')}]")
                else:
                    text = parsed.get("text", "")
                    print(f"  Text (first 120 chars): {text[:120]}")
                error = None
            except Exception as e:
                print(f"  ERROR: {e}")
                raw = ""
                parsed = {"type": "error", "error": str(e)}
                error = str(e)

            results["runs"].append({
                "model_key":    model_key,
                "model":        config["model"],
                "model_type":   config["model_type"],
                "test_id":      test["id"],
                "test_label":   test["label"],
                "prompt":       test["prompt"],
                "expect_note":  test["expect_note"],
                "raw_response": raw,
                "parsed":       parsed,
                "error":        error,
            })

    return results


def main():
    parser = argparse.ArgumentParser(description="Run doit model comparison tests")
    parser.add_argument("--skip-api",   action="store_true", help="Skip the hosted API model")
    parser.add_argument("--skip-local", action="store_true", help="Skip both local Ollama models")
    args = parser.parse_args()

    results = run_all(skip_api=args.skip_api, skip_local=args.skip_local)

    out_file = f"comparison_results_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"\n{'='*60}")
    print(f"Results saved to: {out_file}")
    print("Send that file back along with any notes about timing / failures.")


if __name__ == "__main__":
    main()
