# RDPilot — AI‑Controlled Desktop Agent (Experimental)

**RDPilot** is an experimental, vibe‑coded console app (C# / .NET 9, Windows) that lets a AI operate your desktop by looking at screenshots and emitting actions (keyboard, mouse, scroll, etc.).

* Best results so far with **`gpt-5`**; older models can be **faster**, but are usually less reliable.
* Designed for **Windows 10/11**, .NET **9** is required.


> ⚠️ Currently RDPilot captures and operates only the primary display.

---

## Requirements

* **Windows 10/11**
* **.NET 9** runtime or SDK (**recommended**)
* **.NET 9.0 Desktop Runtime (Windows Desktop Runtime)**
* OpenAI API KEY

---

## What can it do?

Give the model a goal and it will iteratively act on your desktop. For example:

```
open Edge browser, go to Google.com, and search for the term 'life'
```

---

## How it works

**1. Task Retrieval (Prompt)**  
The application first retrieves a **prompt** that defines the goal to be achieved (the task description for the model).  

**2. Initial Screenshot & Model Input**  
A screenshot (PNG) of the primary screen is captured.  
- A **white + red rounded focus ring** (from UI Automation) highlights the element that currently has keyboard focus.  
- An **optional pixel grid overlay** may be added to assist with precise coordinate selection.  

This screenshot, along with the task prompt, is then sent to the LLM.  

**3. Model Decision**  
The LLM (e.g., GPT-5) responds with **exactly one action** to be executed, following a strict JSON schema.  
The available actions include:  
- `paste_text`
- `focus_uia`
- `click_uia`
- `open_url` / `launch_app` (only when explicitly enabled)
- `run_command` (only when explicitly enabled)
- `keys`  
- `type_text`  
- `move`  
- `click`  
- `double_click`  
- `scroll`  
- `request_crop`  
- `point`  
- `aim`  
- `wait`  
- `done`  

**4. Action Execution**  
The application executes the given action via **WinAPI**.  

**5. Iterative Loop**  
After execution, a new screenshot is generated and sent back to the model.  
The model decides the next action.  
This process repeats in a loop until the model returns the `done` action, which signals that the initial task goal (from the prompt) has been achieved.  

> The app writes **logs to files**: screenshots, crops/overlays, and request/response JSONs (see *Output & logs*).

---

## Output & logs

Logs are stored in the following folders:

* **`/screens`**
  Full screenshots (`<id>_<step>.png`), optional **crop** images, **focus\_uia** crops, and **aim\_overlay** images.
* **`/requests`**
  JSON **request/response** payloads per step (`*_request.json`, `*_response.json`) + verifier requests when the model returns `done`.
* **`/logs`**
  A per‑run **console log** (`<id>.log`) that mirrors what you see in the terminal.

---

## Setup

1. Install the **.NET 9** SDK or runtime.
2. Install the **.NET 9** Desktop Runtime (Windows Desktop Runtime).
3. Set your OpenAI API key:

   * PowerShell: `setx OPENAI_API_KEY "sk-..."`
   * Or export in your shell/session.
4. Build or run:

   * Build: `dotnet build -c Release`
   * Run:   `dotnet run --project .`  (or execute your built `.exe`)

---

## Abort Run

**Abort** the current run anytime with **Ctrl+Alt+Q**.

---

## Environment variables & CLI flags

| Purpose                   | Env var                                           | CLI flag                         | Notes |
| ------------------------- | ------------------------------------------------- | -------------------------------- | ----- |
| OpenAI API key            | `OPENAI_API_KEY`                                  | —                                | **Required** |
| Runtime profile           | `RDPILOT_PROFILE=fast/balanced/quality`, `FAST_MODE=1`, `QUALITY_MODE=1` | `--fast` / `--balanced` / `--quality` | Default `fast`; `RDPILOT_PROFILE` overrides mode aliases |
| Model                     | `OPENAI_MODEL=gpt-5.5`                            | `--model <model>`                | Default `gpt-5.5` |
| Q&A model                 | `OPENAI_QA_MODEL=<model>`                         | `--qa-model <model>`             | Falls back to `--model` |
| Verify model              | `OPENAI_VERIFY_MODEL=<model>`                     | `--verify-model <model>`         | Falls back to `--model` |
| Reasoning effort          | `OPENAI_REASONING_EFFORT=low/medium/high/...`    | `--effort <effort>`              | Use `default`, `none`, `minimal`, `low`, `medium`, `high`, or `xhigh`; sent only for reasoning models |
| Q&A reasoning effort      | `OPENAI_QA_REASONING_EFFORT=low/medium/high/...` | `--qa-effort <effort>`           | Default `low`; `default` means API/model default |
| Verify reasoning effort   | `OPENAI_VERIFY_REASONING_EFFORT=low/medium/high/...` | `--verify-effort <effort>`   | Default `low`; `default` means API/model default |
| Mouse actions             | `MOUSE_ENABLED=1/true/yes` or `0/false/no`        | `--mouse` / `--no-mouse`         | Default **on** |
| Post-action UI delay (ms) | `POST_ACTION_DELAY_MS=###`                        | `--delay <ms>`                   | Default `300` in fast mode |
| Pixel grid overlay        | `GRID_STEP_PX=###` or `0`                         | `--grid <px>`                    | Default `0`; e.g. `100` for a 100-px grid |
| Max control steps         | `MAX_STEPS=###`                                   | `--max-steps <n>`                | Default `10000`; hard cap for one control-loop goal |
| Max wait duration         | `MAX_WAIT_SECONDS=###`                            | `--max-wait <seconds>`           | Default `30` in fast mode; `0` disables capping |
| Output token cap          | `MAX_OUTPUT_TOKENS=###`                           | `--max-output-tokens <n>`        | Default `300` in fast mode |
| Q&A output token cap      | `QA_MAX_OUTPUT_TOKENS=###`                        | `--qa-max-output-tokens <n>`     | Default `300` in fast mode |
| Verify output token cap   | `VERIFY_MAX_OUTPUT_TOKENS=###`                    | `--verify-max-output-tokens <n>` | Default `120` in fast mode |
| Incomplete output retry   | `INCOMPLETE_MAX_OUTPUT_RETRIES`, `INCOMPLETE_MAX_OUTPUT_TOKEN_CAP` | `--incomplete-max-output-retries <n>`, `--incomplete-max-output-token-cap <n>` | Retries `status=incomplete` / `max_output_tokens` with a larger cap and lower reasoning effort |
| Max action text length    | `MAX_ACTION_TEXT_CHARS=###`                       | `--max-action-text-chars <n>`    | Default `3000`; long `paste_text` / `type_text` content should be split across actions |
| Q&A screenshot width      | `QA_SCREENSHOT_MAX_WIDTH=###`                     | `--qa-screenshot-max-width <px>` | Default `1024` in fast mode; `0` uses normal send image |
| Verify screenshot width   | `VERIFY_SCREENSHOT_MAX_WIDTH=###`                 | `--verify-screenshot-max-width <px>` | Default `1024` in fast mode; `0` uses normal send image |
| Text verbosity            | `TEXT_VERBOSITY=low/medium/high`                  | `--verbosity <level>`            | Default `low` in fast mode |
| History context           | `HISTORY_TAIL_CHARS=###`                          | `--history-chars <n>`            | Default `1200` in fast mode; `0` disables step history |
| History line limit        | `HISTORY_TAIL_LINES=###`                          | `--history-lines <n>`            | Default `12` in fast mode; keeps only complete recent history entries |
| Stagnation guard          | `MAX_STAGNATION_STEPS=###`                        | `--max-stagnation <n>`           | Default `8` in fast mode; `0` disables |
| Repeated action guard     | `MAX_REPEATED_ACTIONS=###`                        | `--max-repeated-actions <n>`     | Default `5` in fast mode; `0` disables |
| Repeat cooldown           | `ACTION_REPEAT_COOLDOWN_STEPS=###`                | `--repeat-cooldown <n>`          | Default `2` in fast mode; temporarily blocks identical ineffective UI actions |
| Model failure guard       | `MAX_MODEL_FAILURES=###`                          | `--max-model-failures <n>`       | Default `2` in fast mode; keeps transient API failures from aborting immediately |
| Local action failure guard | `MAX_ACTION_FAILURES=###`                        | `--max-action-failures <n>`      | Default `2` in fast mode; feeds executor failures back to the model before aborting |
| Screen polling            | `SCREEN_POLLING=1/0`                              | `--screen-polling` / `--no-screen-polling` | Default on; waits for visual stability instead of only fixed sleeps |
| Screen poll timing        | `SCREEN_POLL_INITIAL_DELAY_MS`, `SCREEN_POLL_INTERVAL_MS`, `SCREEN_POLL_TIMEOUT_MS` | `--screen-poll-initial-delay <ms>`, `--screen-poll-interval <ms>`, `--screen-poll-timeout <ms>` | Defaults `120/150/1200` in fast mode |
| Extra unchanged wait      | `WAIT_NO_CHANGE_EXTRA_MS=###`                     | `--wait-no-change-extra <ms>`    | Default `750`; used when a `wait` action leaves the screen unchanged |
| Screenshot sanity checks  | `SCREEN_SANITY_CHECKS=1/0`                        | `--screen-sanity` / `--no-screen-sanity` | Warns once per run for black/uniform screens, tiny resolution, or the RDPilot console in front |
| SendInput retry           | `SENDINPUT_MAX_RETRIES`, `SENDINPUT_RETRY_DELAY_MS` | `--sendinput-retries <n>`, `--sendinput-retry-delay <ms>` | Defaults `2` and `30`; retries transient keyboard/mouse injection failures |
| Adaptive effort           | `ADAPTIVE_REASONING_EFFORT=1/0`                   | `--adaptive-effort` / `--no-adaptive-effort` | Default on; disable to keep control effort fixed |
| Screenshot max width      | `SCREENSHOT_MAX_WIDTH=###`                        | `--screenshot-max-width <px>`    | Default `1280`; `0` keeps original |
| Screenshot format         | `SCREENSHOT_FORMAT=jpeg/png`                      | `--screenshot-format <format>`   | Default `jpeg` in fast mode |
| Focused overview width    | `FOCUSED_OVERVIEW_MAX_WIDTH=###`                  | `--focused-overview-max-width <px>` | Default `640` in fast mode; used for the full-screen overview when a focus crop is sent |
| Crop max width            | `CROP_MAX_WIDTH=###`                              | `--crop-max-width <px>`          | Default `768` in fast mode; affects `aim`/focus crops sent to the model |
| Crop format               | `CROP_FORMAT=jpeg/png`                            | `--crop-format <format>`         | Default `jpeg` in fast/balanced, `png` in quality |
| JPEG quality              | `SCREENSHOT_JPEG_QUALITY=1..100`                  | `--jpeg-quality <n>`             | Default `80` in fast mode |
| Screenshot log format     | `SCREEN_LOG_FORMAT=jpeg/png`                      | `--screen-log-format <format>`   | Default `jpeg` in fast/balanced, `png` in quality |
| Screenshot log max width  | `SCREEN_LOG_MAX_WIDTH=###`                        | `--screen-log-max-width <px>`    | Default `1280` in fast mode; `0` keeps original |
| Focus UIA metadata        | `INCLUDE_FOCUS_UIA=1/0`                           | `--focus-uia` / `--no-focus-uia` | Default on; controls focused-element rect/summary/overlay |
| Focus UIA crop            | `INCLUDE_FOCUS_UIA_CROP=1/0`                      | `--debug-images` enables it      | Default off in fast mode |
| Verify mode               | `VERIFY_MODE=auto/always/off`                     | `--verify <mode>`                | Default `auto` |
| Auto verify early steps   | `VERIFY_EARLY_STEPS=###`                          | `--verify-early-steps <n>`       | Default `2`; set `0` to verify only high-impact goals in `auto` mode |
| Verify low confidence     | `VERIFY_LOW_CONFIDENCE_THRESHOLD=0..1`            | `--verify-low-confidence <n>`    | Default `0.75`; in `auto`, low-confidence `done` still triggers verifier after early-step window |
| Skip verify confidence    | `SKIP_VERIFY_CONFIDENCE_THRESHOLD=0..1`           | `--skip-verify-confidence <n>`   | Default `0.92`; non-critical high-confidence `done` can skip verifier after early-step window |
| Verify screenshot refresh | `REFRESH_SCREENSHOT_BEFORE_VERIFY=1/0`            | `--refresh-before-verify` / `--no-refresh-before-verify` | Default off; verifier reuses the current screenshot |
| OpenAI retries            | `OPENAI_MAX_RETRIES=###`                          | `--retries <n>`                  | Default `2` |
| OpenAI timeout            | `OPENAI_TIMEOUT_SECONDS=###`                      | `--openai-timeout <seconds>`     | Default `600`; `0` disables the client-side timeout |
| Prompt cache              | `PROMPT_CACHE=1/0`                                | `--prompt-cache` / `--no-prompt-cache` | Default on |
| Prompt cache key          | `PROMPT_CACHE_KEY=<key>`                          | `--prompt-cache-key <key>`       | Default `rdpilot-control-v1`; RDPilot appends `:control`, `:qa`, or `:verify` |
| Previous response state   | `USE_PREVIOUS_RESPONSE_ID=1/0`                    | `--previous-response-state` / `--no-previous-response-state` | Default off; opt-in Responses API state between control steps |
| Omit unchanged screen     | `OMIT_UNCHANGED_SCREEN_IMAGE=1/0`                 | `--omit-unchanged-screen` / `--no-omit-unchanged-screen` | Default off; only applies with previous-response state and unchanged screen fingerprint |
| Long text paste threshold | `CLIPBOARD_PASTE_THRESHOLD=###`                   | `--paste-threshold <n>`          | Default `120`; long `type_text` uses clipboard |
| AIM/focus crop size       | `FOCUS_CROP_SIZE=###`                             | `--focus-crop-size <px>`         | Default `320`; reduce for smaller crop payloads |
| Request logs              | `LOG_REQUESTS=1/0`                                | `--no-request-logs`              | Default on |
| Request log formatting    | `PRETTY_REQUEST_LOGS=1/0`                         | `--pretty-request-logs` / `--compact-request-logs` | Default compact in fast/balanced, pretty in quality |
| Screenshot logs           | `LOG_SCREENS=1/0`                                 | `--no-screen-logs`               | Default on; disabling skips PNG screenshot artifacts while requests still use in-memory images |
| Real UI only              | `REAL_UI_ONLY=1/0`                                | `--real-ui-only`                 | Forces local adapters off even if config/env enabled them |
| High-level local actions  | `ALLOW_HIGH_LEVEL_ACTIONS=1/0`                    | `--allow-high-level-actions`     | Default off; real UI is the default |
| Run commands              | `ALLOW_RUN_COMMAND=1/0`                           | `--allow-run-command`            | Default off |
| Batch response candidates | `EXECUTE_MULTI_ACTION_CANDIDATES=1/0`             | `--batch-candidates` / `--no-batch-candidates` | Default off; experimental, UI-safe follow-up actions only |
| Console auto-hide         | `AUTO_HIDE_CONSOLE=1/0`                           | `--auto-hide-console` / `--no-auto-hide-console` | Default on; hides RDPilot console before screenshots/control actions |
| Console placement         | `MINIMIZE_CONSOLE_DURING_RUN=1/0`                 | `--minimize-console`             | Compatibility flag that minimizes instead of hiding before the run |
| Console restore           | `RESTORE_CONSOLE_AFTER_RUN=1/0`                   | `--restore-console` / `--no-restore-console` | Default on |
| UIA target list           | `INCLUDE_UIA_TARGETS=1/0`                         | `--max-uia-targets <n>`          | Sends indexed UI Automation targets to the model |
| UIA target name length    | `UIA_TARGET_NAME_MAX_CHARS=###`                   | `--uia-name-chars <n>`           | Default `48` in fast mode; trims long labels in prompts |
| UIA summary length        | `UIA_SUMMARY_MAX_CHARS=###`                       | `--uia-summary-chars <n>`        | Default `320` in fast mode |
| UIA scan budget           | `UIA_SCAN_TIME_BUDGET_MS=###`                     | `--uia-scan-ms <ms>`             | Default `60` in fast mode; keeps UI metadata collection bounded |
| UIA node budget           | `MAX_UIA_NODES_SCANNED=###`                       | `--max-uia-nodes <n>`            | Default `400` in fast mode |
| UIA candidate multiplier  | `UIA_CANDIDATE_MULTIPLIER=###`                    | `--uia-candidate-multiplier <n>` | Default `4` in fast mode; gathers extra candidates before ranking |
| UIA max area ratio        | `UIA_MAX_AREA_RATIO=0..1`                         | `--uia-max-area-ratio <n>`       | Default `0.45` in fast mode; filters large low-value containers |
| UIA target reuse          | `REUSE_UIA_TARGETS_ON_NO_CHANGE=1/0`              | `--reuse-uia-targets` / `--no-reuse-uia-targets` | Default on; skips a fresh UIA scan when the screen fingerprint did not change |
| Artifact retention        | `MAX_ARTIFACTS_PER_DIR=###`                       | `--max-artifacts <n>`            | Default `500`; keeps recent run groups per artifact folder; `0` disables cleanup |
| Effective config          | —                                                 | `--print-config`                 | Prints the resolved config without requiring an API key |
| Log analysis              | —                                                 | `--analyze-logs`                 | Summarizes saved responses, screenshots, and logs |
| Response replay           | —                                                 | `--replay-response <path>`       | Parses a saved response and shows safe follow-up candidates |
| Request replay            | —                                                 | `--replay-request <path>` / `--replay-request-dry-run` | Rehydrates saved request logs from `file://` images; dry-run validates without API key |

---

## Config File

RDPilot also reads `rdpilot.json` from the working directory or executable directory. Environment variables and CLI flags override it. Profile flags are applied as a base first, so explicit CLI settings such as `--history-chars` or `--uia-scan-ms` override the selected profile regardless of argument order.

Example:

```json
{
  "profile": "fast",
  "model": "gpt-5.5",
  "qaModel": "gpt-5-mini",
  "verifyModel": "gpt-5-mini",
  "qaReasoningEffort": "low",
  "verifyReasoningEffort": "low",
  "maxSteps": 10000,
  "maxWaitSeconds": 30,
  "historyTailChars": 1200,
  "historyTailLines": 12,
  "qaMaxOutputTokens": 300,
  "verifyMaxOutputTokens": 120,
  "incompleteMaxOutputRetries": 2,
  "incompleteMaxOutputTokenCap": 4096,
  "maxActionTextChars": 3000,
  "qaScreenshotMaxWidth": 1024,
  "verifyScreenshotMaxWidth": 1024,
  "adaptiveReasoningEffort": true,
  "maxStagnationSteps": 8,
  "maxRepeatedActions": 5,
  "actionRepeatCooldownSteps": 2,
  "maxModelFailures": 2,
  "maxActionFailures": 2,
  "screenPolling": true,
  "screenPollInitialDelayMs": 120,
  "screenPollIntervalMs": 150,
  "screenPollTimeoutMs": 1200,
  "waitNoChangeExtraMs": 750,
  "screenSanityChecks": true,
  "sendInputMaxRetries": 2,
  "sendInputRetryDelayMs": 30,
  "screenshotMaxWidth": 1280,
  "screenshotFormat": "jpeg",
  "focusedOverviewMaxWidth": 640,
  "cropMaxWidth": 768,
  "cropFormat": "jpeg",
  "screenLogFormat": "jpeg",
  "screenLogMaxWidth": 1280,
  "includeFocusUia": true,
  "verifyMode": "auto",
  "verifyEarlySteps": 2,
  "verifyLowConfidenceThreshold": 0.75,
  "skipVerifyConfidenceThreshold": 0.92,
  "refreshScreenshotBeforeVerify": false,
  "openAiTimeoutSeconds": 600,
  "focusCropSize": 320,
  "promptCache": true,
  "promptCacheKey": "rdpilot-control-v1",
  "usePreviousResponseId": false,
  "omitUnchangedScreenImage": false,
  "realUiOnly": false,
  "allowHighLevelActions": false,
  "autoHideConsoleDuringRun": true,
  "restoreConsoleAfterRun": true,
  "prettyRequestLogs": false,
  "executeMultiActionCandidates": false,
  "includeUiaTargets": true,
  "maxUiaTargets": 20,
  "uiaTargetNameMaxChars": 48,
  "uiaSummaryMaxChars": 320,
  "uiaScanTimeBudgetMs": 60,
  "maxUiaNodesScanned": 400,
  "uiaCandidateMultiplier": 4,
  "uiaMaxAreaRatio": 0.45,
  "reuseUiaTargetsWhenScreenUnchanged": true
}
```

---

## Profiles

* `fast`: `effort=low`, JPEG screenshots downscaled to 1280px, 640px full-screen overview when a focus crop is sent, JPEG crops up to 768px, JPEG screen logs up to 1280px, short output, short step history, adaptive verify, no debug overlays.
* `balanced`: JPEG screenshots up to 1600px, JPEG crops up to 1024px, JPEG screen logs up to 1600px, focus UIA crop enabled, short output, medium step history.
* `quality`: original PNG screenshots and crops, original PNG screen logs, focus UIA crop, debug overlays, longer step history, verifier always on, `effort=medium`.

The control loop can temporarily raise reasoning effort to `medium`/`high` when it detects stagnation or repeated ineffective actions.
Use `--no-adaptive-effort` to keep GPT-5.5 on the configured effort for latency-sensitive runs.
Q&A and verifier calls can use separate effort settings, so the main control loop can spend more reasoning only when needed while cheaper helper calls stay on `low`. For helper calls, `default` is explicit: it omits `reasoning.effort` instead of falling back to the control-loop effort.
Verifier calls use a smaller output-token cap by default because they return only `yes`/`no` plus a short reason.
Q&A and verifier calls can use their own smaller screenshot width, keeping helper calls cheaper than the main control loop.
Verifier prompts keep `SCREEN_SIZE` aligned with the actual helper image after downscaling.
In `auto` verify mode, the early-step verification window is configurable with `--verify-early-steps`.
Control actions include a `confidence` value. After the early verification window and outside high-impact goals, low-confidence `done` actions still trigger the verifier while high-confidence simple completions can finish without an extra model call.
The high-confidence skip threshold is configurable with `--skip-verify-confidence`, while high-impact goals still verify in `auto` mode.
The control system prompt avoids per-step action-list churn; dynamic allowed actions stay in the JSON schema, which improves prompt-cache reuse across steps.
The configured prompt-cache key is treated as a base key; control, Q&A, and verifier calls use separate scoped keys so different prompt prefixes do not compete with each other.
Control requests split stable user context from dynamic history/metadata, letting prompt caching reuse more of the prefix between steps.
`--previous-response-state` enables Responses API `previous_response_id` for control steps and omits the explicit action-history tail after the first turn, while still sending fresh screenshots and current metadata.
With `--omit-unchanged-screen`, unchanged-screen turns using previous-response state can skip the full-screen image and send only current metadata plus any active crop images.

By default RDPilot uses the real visible UI: keyboard, mouse, clipboard paste, screenshots, and UI Automation metadata. High-level local shortcuts such as `open_url`, `launch_app`, and `run_command` are disabled unless you explicitly enable them.
Use `--real-ui-only` to force those local adapters back off even when an old config file or environment variable enabled them.
When mouse mode is disabled or the current step has no indexed UIA targets, dependent actions are removed from the action schema instead of being offered to the model and rejected later.
The control prompt is also conditional: mouse/aim/click rules are omitted when mouse mode is off, while keyboard and clipboard guidance remains available.

UI Automation target discovery is bounded by time and node count. This avoids slow full-tree scans in complex apps while still giving the model real UI targets such as buttons, inputs, lists, and menu items.
UIA target names and focused-element summaries are trimmed in fast mode to reduce prompt size; increase `--uia-name-chars` or `--uia-summary-chars` when an app uses long, meaningful labels.
RDPilot collects a small over-sampled UIA candidate pool, deduplicates it, filters large low-value containers, and ranks actionable controls before sending indexed targets to the model. This keeps the prompt closer to real clickable UI instead of noisy window/pane structure.
Control and Q&A prompts include the active window title, foreground process name, and active-window geometry/visibility state, giving the model a cheap state probe without adding another screenshot or high-level adapter.
If the target window is too small, clipped, or partly off-screen, the control prompt tells the model to correct it through the real UI, usually with keyboard window-management shortcuts such as `Win+Up`; RDPilot does not automatically move or maximize target windows.
Verifier prompts use the same active-window, foreground-process, active-window-geometry, focused-UIA, and blocking-dialog hints as control/Q&A calls.
Use `--no-focus-uia` together with `--max-uia-targets 0` for workflows where screenshot-only control is enough and all UIA calls should be minimized.
Request logging reuses the UIA target list gathered for the actual model request instead of scanning the UI tree a second time.
The control request itself also reuses the prebuilt UIA target list for schema generation and prompt text, avoiding a duplicate scan in changed-screen steps.
UIA target scanning reuses the dimensions from the current screenshot instead of querying screen geometry again.
Focused UIA rectangle and summary are captured from a single focused-element snapshot per screenshot instead of two separate focused-element calls.
The active window title and focused UIA summary are captured once per step and reused for both the request and its log entry.
When request logging is disabled, RDPilot skips constructing the redacted log payloads instead of building and discarding them.
Request logs are compact by default in fast/balanced profiles to reduce disk I/O; use `--pretty-request-logs` when inspecting them by hand.
Request log JSON is serialized directly to UTF-8 bytes, avoiding an intermediate string allocation.
Artifact cleanup keeps recent run groups together instead of deleting arbitrary individual screenshots or request files.
Step history is maintained only in a bounded tail buffer instead of retaining the full list and rebuilding `string.Join` over the whole run on every model call.
The history tail also keeps only complete recent entries, capped by `--history-lines`, so old partial lines do not consume prompt tokens.
Debug AIM overlays are scaled to the saved screen-log image size, so overlays remain accurate even when screen logs are downscaled.
Run metrics include request/response byte counts plus screenshot, image-encoding, screen-log, and UIA timings, which makes it easier to see whether screenshot compression and UIA prompt trimming are actually reducing payload and local overhead.
Run metrics also include input, cached, output, and reasoning token counts for immediate feedback on prompt-cache behavior.
If prompt caching is enabled but a multi-call run reports zero cached tokens, RDPilot prints a prompt-cache warning at the end of the run.
Run metrics also include OpenAI retry count, local action execution timing, lightweight screen-probe timing, screenshot sanity warnings, and request/response artifact log write timing, so slow runs can separate API latency from local overhead.
When the full screenshot send profile matches the screen-log profile, RDPilot encodes it once and reuses the encoded bytes for both the API payload and the log file.
The same encoded-byte reuse is applied to crop images when crop-send and screen-log profiles produce identical output.
OpenAI request bodies are serialized directly to UTF-8 bytes and sent as `ByteArrayContent`, avoiding an extra string-to-bytes pass.
Pixel coordinates, bounding boxes, UIA indexes, and wait durations are bounded in the JSON schema for each current screen, reducing invalid model actions and retry loops.
When a screenshot is downscaled before sending, `SCREEN_SIZE` and the action schema use the downscaled image dimensions seen by the model; RDPilot maps those coordinates back to the real desktop before moving or clicking the mouse.
Keyboard action arrays are bounded in the schema so the model keeps shortcuts compact instead of emitting long key scripts.
Mouse coordinates are clamped locally as a second line of defense for manual JSON, replayed responses, and older logs.
Scroll actions are bounded and use the documented convention that positive `scroll_dy` moves the visible page down.
Additional crops from `aim`, `request_crop`, and focus UIA use their own send profile, so fast mode can keep crop payloads bounded without changing full-screenshot behavior.
When an `aim` or `request_crop` focus crop is present, the crop is sent before the full-screen image and the full-screen image is reduced to a small overview in fast/balanced profiles.
The square crop size used for point-based `aim`/`request_crop` is configurable with `--focus-crop-size`.
Local observation actions (`aim`, `point`, `request_crop`) skip the global post-action UI delay because they do not mutate the target app; plain cursor moves use only their short hover delay.
Screen-change fingerprints are computed from locked bitmap memory instead of per-pixel GDI calls, reducing local overhead in long control loops.
Screen-change fingerprints are computed before RDPilot draws focus/grid overlays, so polling compares clean desktop captures against clean desktop probes.
Fingerprint downsampling uses a fast resize path; high-quality resize is reserved for images sent to the model or saved for inspection.
Image data URLs are built from the encoding buffer without an extra `ToArray()` copy, reducing allocations for screenshot-heavy runs.
Screenshot encoding uses the original bitmap directly when no downscale is needed, avoiding a full-size bitmap copy in quality/original-width modes.
Short Unicode typing is batched into fewer `SendInput` calls; long text still uses clipboard paste based on `--paste-threshold`.
Clipboard paste retries briefly when the clipboard is busy, avoiding failed long-text actions caused by transient ownership conflicts.
Keyboard and mouse `SendInput` calls retry briefly on transient Windows injection failures before aborting the action.
Keyboard shortcuts are batched into a single virtual-key `SendInput` sequence when possible, while unusual text keys keep the compatibility fallback.
Simple key sequences such as repeated `tab` plus `enter` are also batched into one `SendInput` call when all keys map to virtual keys.
Common key-name aliases such as `pgdn`, `del`, `ins`, and `arrowleft` are accepted to avoid aborting a run over naming differences.

Loop guards stop runs that keep sending expensive model calls without visible progress. Set `--max-stagnation 0` or `--max-repeated-actions 0` for long workflows that are expected to look unchanged for many steps. Long `wait` actions are capped by `--max-wait`, and the action schema advertises that cap to the model.
Repeated-action detection distinguishes different `request_crop`/`point` regions, so useful visual refinement is not mistaken for the same ineffective action.
Mouse clicks and double-clicks that land in the same small screen region are also clustered for repeat detection, so tiny coordinate changes do not let the model keep retrying an ineffective click.
Local observation actions are not counted as UI stagnation, avoiding unnecessary adaptive effort escalation after `aim` or crop refinement.
When the verifier rejects `done`, its reason is promoted into the next control prompt as `LAST_VERIFY_REJECTION` instead of being buried only in the action history.
When the local delta/repeat guard sees no visible progress, the next prompt includes an explicit strategy hint telling the model not to repeat the same action.
Identical ineffective UI actions can also be put on a short local cooldown, which prevents immediate repeat loops while still allowing a different visible-UI route.
After a click with little or no visible progress, the next prompt includes a precision hint telling the model not to repeat the same point and to use crop/aim or keyboard expansion for tiny controls, tree expanders, lists, and menus. The hint remains available briefly if an intermediate keyboard attempt also produces no visible progress.
When local action execution fails, RDPilot records `LAST_EXECUTOR_FAILURE` in the next prompt and gives the model a chance to choose a different action before the action-failure guard stops the run.
After mutating UI actions, RDPilot can poll lightweight screen fingerprints until the screen is stable instead of relying only on a fixed post-action sleep. This keeps quick UI steps snappy while giving slower transitions time to settle before the next expensive model call.
If a long `wait` action leaves the screen visually unchanged, RDPilot waits a short extra interval before asking the model again.
Screenshot sanity checks warn when the captured screen is nearly black, nearly uniform, unexpectedly small, or when the RDPilot console itself is the foreground window and may be covering the target app.
Active-window and focused-UIA metadata are scanned for modal, permission, and UAC hints; when detected, the next prompt tells the model to resolve the visible dialog explicitly.
By default RDPilot minimizes its own console before control-loop screenshots and actions, then restores it when the run finishes. This prevents the console from becoming the UI target or covering the app being controlled.
Ctrl+Alt+Q cancels in-flight OpenAI HTTP calls and retry backoff, so aborting a slow GPT-5.5 response no longer waits for the request timeout.
The same abort token cancels long `wait`, verifier settle delay, batched waits, and post-action delays.
After retryable OpenAI failures such as 5xx, timeout, or transport errors, the control loop keeps the goal alive for a small number of attempts instead of aborting after the first failed call. Non-retryable errors and parse errors still stop the goal.
If a response completes as `status=incomplete` because `max_output_tokens` was spent before valid JSON was emitted, RDPilot retries with a larger output cap and `reasoning.effort=low`.
Very long `paste_text` and `type_text` payloads are capped by schema so the model splits large content across multiple real-UI paste actions instead of emitting an oversized JSON response that can be truncated.
If `paste_text` or `type_text` produces no visible screen change, the next prompts include `TEXT_INPUT_HINT`, telling the model to fix focus/editability or choose another visible UI route before repeating text input.
After repeated no-change text-input attempts, RDPilot temporarily blocks further `paste_text`, `type_text`, and paste shortcuts so the model must establish focus/editability or choose another visible UI path first.

`--batch-candidates` is intentionally disabled by default. In existing logs, extra JSON candidates can be alternate model answers rather than a safe sequence of next steps, so this mode should be tested per workflow before regular use.
Duplicate parsed response candidates are ignored before warning or optional batching, reducing noise from APIs that expose the same JSON in multiple response fields.
Parsed action candidates must contain a known action `type`, so Q&A/verifier payloads are not miscounted as control actions in replay or analysis.
Optional batched follow-ups also skip duplicate action signatures, avoiding repeated text/keyboard actions from alternate candidates.
`--analyze-logs` uses the same parsed-candidate deduplication, so multi-action counts better reflect genuine alternate actions.

Use `--analyze-logs` to inspect previous runs and find slow calls, runtime metrics, request payload size, prompt text volume, model/token distribution, multi-action responses, largest screenshots, rejected verifier decisions, and HTTP errors.
The analyzer reports input-token and cached-token totals, making prompt-cache hit rate visible after profile or prompt changes.
Use `--replay-request <request.json> --replay-request-dry-run` to validate that a saved request can be reconstructed from its screenshot artifacts without touching the desktop. Without dry-run, RDPilot resends the hydrated request to OpenAI and writes a sibling `_replay_response.json`, which helps compare prompt/model/effort changes without operating the real UI.

---

## Examples

### Task (control loop)

```
open Edge browser, go to Google.com, and search for the term 'life'
```

### Faster GPT-5.5 run

```
dotnet run --project RDPilot -- --effort low "open Edge browser, go to Google.com, and search for the term 'life'"
```

### Q\&A (screenshot analysis)

```
/ask where do you see the Edge app icon?
```

---

## Safety, limitations & disclaimer

This is an **experimental** project built for exploration and learning. It simulates real input on your machine and may act unpredictably (mis‑clicks, wrong targets, etc.).

* Use on a **throwaway VM** or a non‑critical environment when possible.
* You run it **at your own risk**. No warranty of any kind.

---

## License

MIT.

If you improve this code, I’ll be happy to accept the changes in a PR :)

