# Gemini Model Auto-Fallback Chain

**Date:** 2026-04-16
**Status:** Approved by user; not yet implemented
**Trigger for this feature:** On 2026-04-16 `gemini-2.5-flash` returned
`503 This model is currently experiencing high demand` repeatedly on a real
PDF run, blocking the user for ~18 minutes (Google SDK inner retry deadline
is 600 s per call). User wants automatic failover so a single-model outage
doesn't stall the whole workflow.

## Goal

When the primary model is busy or unavailable, automatically try the next
model in a chain so the user's PDF run completes without manual
intervention.

## Fallback chain

```
gemini-2.5-flash   →   gemini-2.0-flash   →   gemini-2.5-pro
(primary, cheap)       (separate pool,       (expensive last
                        cheap)                resort)
```

Rationale:
- **2.5-flash first:** user's chosen default, cheapest ($0.075 / 1M input).
- **2.0-flash second:** different capacity pool, not affected by 2.5 spikes
  in most outages. Still cheap (~$0.10 / 1M input). OCR quality on
  handwritten notebooks is close enough to 2.5-flash.
- **2.5-pro last:** ~17× more expensive (~$1.25 / 1M input) but the most
  reliable. Only reached if both flash variants are out.

## Trigger rules

Fall back **immediately** (don't waste retries on the same model) when:
- `google.api_core.exceptions.ServiceUnavailable` (503, "high demand")
- `google.api_core.exceptions.ResourceExhausted` (429, quota / rate limit)
- After all 3 retries of `DeadlineExceeded` on the current model (a chronic
  timeout — different model might respond).

Do NOT fall back on:
- `InvalidArgument` (bad prompt — same in other models, pointless retry)
- `PermissionDenied` (auth issue — same everywhere)
- JSON parse errors (model-specific output shape — may actually be
  different across models, but safer to surface so we can debug)

## Where to implement

`reader.py` → `_call_gemini_on_image`.

**Current signature:**
```python
def _call_gemini_on_image(model, image: Image.Image) -> dict:
```

**New signature:**
```python
def _call_gemini_on_image(model_name_chain: list[str], image: Image.Image) -> dict:
```

Model instances are created inside this function (one per attempted
model in the chain). Caller (`read_notebook`) passes the chain, not a
pre-built model instance. This keeps the fallback logic localized.

## Algorithm

```python
def _call_gemini_on_image(model_chain, image):
    from google.api_core import exceptions as gax
    last_err = None
    for model_name in model_chain:
        model = genai.GenerativeModel(model_name)
        for attempt in range(3):
            try:
                _log(f"  [{model_name}] attempt {attempt+1}: sending image ({image.size})")
                t0 = time.time()
                resp = model.generate_content(
                    [PROMPT, image],
                    generation_config={
                        "response_mime_type": "application/json",
                        "temperature": 0.0,
                    },
                    request_options={"timeout": 300},
                )
                _log(f"  [{model_name}] got response in {time.time()-t0:.1f}s")
                return _parse_response(resp)
            except (gax.ServiceUnavailable, gax.ResourceExhausted) as e:
                last_err = e
                _log(f"  [{model_name}] busy ({type(e).__name__}); falling over to next model")
                break  # Don't retry same busy model; go to next in chain.
            except gax.DeadlineExceeded as e:
                last_err = e
                if attempt < 2:
                    time.sleep(2 ** attempt)
                    continue
                _log(f"  [{model_name}] timed out 3×; falling over to next model")
                break  # After 3 timeouts, try next model.
    raise RuntimeError(f"All models in chain failed. Last error: {last_err}")
```

## Call-site change

In `read_notebook`, replace:
```python
model = genai.GenerativeModel(model_name)
...
data = _call_gemini_on_image(model, image)
```
with:
```python
model_chain = [model_name, "gemini-2.0-flash", "gemini-2.5-pro"]
# Dedupe in case user already picked one of the fallbacks as their primary.
seen = set()
model_chain = [m for m in model_chain if not (m in seen or seen.add(m))]
...
data = _call_gemini_on_image(model_chain, image)
```

The threading pool from the parallel-pages change is unaffected — each
worker just receives the chain instead of a model instance.

## Logging

Every fallover writes one line to `gemini_debug.log`:
```
[17:55:50]   [gemini-2.5-flash] busy (ServiceUnavailable); falling over to next model
[17:55:50]   [gemini-2.0-flash] attempt 1: sending image ((2500, 1762))
```

## UI (optional, deferred)

Not in this pass — silent fallback is fine. Later: show the model that was
actually used per page in the preview window (e.g. pill on cube card:
"via pro"), so user sees when they got charged for pro.

## Settings

No new settings. The chain is hard-coded. If the user's chosen model is
ALREADY `gemini-2.5-pro`, the chain just contains pro (and fallbacks are
effectively disabled). For a future iteration we can surface the chain in
Settings, but YAGNI for now.

## Tests / acceptance

1. Run real 7days.pdf when Google is NOT having 503 issues — all pages
   use `gemini-2.5-flash`, no fallback triggered, log shows only
   `[gemini-2.5-flash]` lines. **Critical regression test.**
2. Force-simulate 503 (monkey-patch `generate_content` to raise
   `ServiceUnavailable` the first time) → verify fallover to 2.0-flash
   happens on the first error, no retry loop on flash.
3. Force both flashes to fail → verify pro is attempted; if pro
   succeeds, JSON parses and pipeline continues.
4. Force all three to fail → RuntimeError surfaces cleanly in the
   PreviewWindow (the messagebox error dialog fix from 2026-04-16 is
   already in place).

## Cost guardrails

- Only pro costs ~17× more. Real-world impact: user processes maybe 5-20
  PDFs per day. If 503 happens on 1 PDF and all 5 pages fall through to
  pro, that's ~$0.25 instead of ~$0.015 for that one PDF — acceptable
  for "problemsiz çalışsın" requirement.
- Document this in the README or inline help so user knows the cost
  model.

## Related prior work (already shipped 2026-04-16)

- **MAX_SIDE=2500 cap** on images (reader.py:135) — already built into
  new exe.
- **Parallel page processing** (`ThreadPoolExecutor(max_workers=5)` in
  `read_notebook`) — already coded, not yet validated end-to-end because
  of the 503 outage that prompted this spec.
- **Error dialog fix** (main.py:1896 — traceback written to
  `gemini_debug.log`, messagebox shows short message only) — already
  applied.

## How to resume next session

1. Open new terminal in `C:\Users\Yafka\Desktop\CubeLogReader`.
2. Start Claude Code.
3. Say: **"docs\superpowers\specs\2026-04-16-model-fallback-design.md
   dosyasındaki planı uygula ve sonra exe'yi yeniden build et"**
4. Claude will read this file, implement the changes in reader.py, test
   with `python main.py`, then rebuild via `build_exe.bat`.
