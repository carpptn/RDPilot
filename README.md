# RDPilot — AI‑Controlled Desktop Agent (Experimental)

**RDPilot** is an experimental, vibe‑coded console app (C# / .NET 9, Windows) that lets a Large Language Model (LLM) operate your desktop by looking at screenshots and emitting single, atomic actions (keyboard, mouse, scroll, etc.).

* Best results so far with **`gpt-5`**; older models can be **faster**, but are usually less reliable.
* Designed for **Windows 10/11 (x64)**. .NET **9** is recommended.

> ⚠️ By default the **mouse is disabled** because current models still struggle with precise targeting of on‑screen elements. Keyboard‑first strategies are encouraged. You can enable the mouse with a flag or env var (see below).

> ⚠️ BCurrently RDPilot captures and operates only the primary display (primary monitor).

---

## What can it do?

Give the model a high‑level goal and it will iteratively act on your desktop. For example:

```
open Edge, go to Google.com, and search for the term 'życie'
```

It also supports “Q\&A on screenshot” via `/ask`, e.g.:

```
/ask where do you see the Edge app icon?
```

---

## How it works (high level)

1. **Capture** the primary screen as a PNG. A white+red rounded focus ring (from UI Automation) overlays the element that currently has keyboard focus. An optional pixel **grid overlay** can assist with precise coordinates.
2. **Prompt & call** the LLM (default `gpt-5`) with strict JSON schema: the model must return **exactly one** action per round (`keys`, `type_text`, `move`, `click`, `double_click`, `scroll`, `request_crop`, `point`, `aim`, `wait`, `done`).
3. **Policy gates** are enforced:

   * **Aim‑before‑click**: a `click/double_click` is ignored unless an `aim` defined the target region first. Clicks must provide **explicit coordinates** (`x/y` or `x_px/y_px`) that fall **inside** the active AIM.
   * **AIM expiration**: after a large visual change (`LAST_STEP_DELTA > AimExpireDelta`), the previous AIM is invalidated and must be set again.
   * Prefer deterministic, **keyboard‑first** strategies. Mouse may be disabled.
4. **Execute** via WinAPI (SendInput), wait a short **UI settle delay**, then loop until the goal is reached or step limits are hit.

> The app writes **logs to files**: screenshots, crops/overlays, and request/response JSONs (see *Output & logs*).

---

## Output & logs

All artifacts are stored next to the executable:

* **`/screens`**
  Full screenshots (`<id>_<step>.png`), optional **crop** images, **focus\_uia** crops, and **aim\_overlay** images.
* **`/requests`**
  JSON **request/response** payloads per step (`*_request.json`, `*_response.json`) + verifier requests when the model returns `done`.
* **`/logs`**
  A per‑run **console log** (`<id>.log`) that mirrors what you see in the terminal.

---

## Requirements

* **Windows 10/11 (x64)**
* **.NET 9** runtime or SDK (**recommended**)
* Network access to `api.openai.com`

---

## Setup

1. Install the **.NET 9** SDK or runtime.
2. Set your OpenAI API key:

   * PowerShell: `setx OPENAI_API_KEY "sk-..."`
   * Or export in your shell/session.
3. Build or run:

   * Build: `dotnet build -c Release`
   * Run:   `dotnet run --project .`  (or execute your built `.exe`)

---

## Running

You can pass a **single goal** as an argument, or start with no args to use **interactive mode**.

```bash
# one‑shot goal
RDPilot.exe "open Edge, go to Google.com, and search for the term 'życie'"

# interactive mode
RDPilot.exe
# then type goals or Q&A like:
#   open Edge, go to Google.com, and search for the term 'życie'
#   /ask where do you see the Edge app icon?
```

**Abort** the current run anytime with **Ctrl+Alt+Q** (when the console has focus).

---

## Environment variables & CLI flags

| Purpose                    | Env var                    | CLI flag          | Notes                                  |                                                     |
| -------------------------- | -------------------------- | ----------------- | -------------------------------------- | --------------------------------------------------- |
| OpenAI API key             | `OPENAI_API_KEY`           | —                 | **Required**                           |                                                     |
| Enable mouse actions       | `MOUSE_ENABLED=1/true/yes` | `--mouse`         | Default **off** (aiming is unreliable) |                                                     |
| Post‑action UI delay (ms)  | `POST_ACTION_DELAY_MS=###` | `--delay <ms>`    | Default `1000`                         |                                                     |
| Pixel grid overlay         | `GRID_STEP_PX=###` or `0`  | \`--grid \<px     | off>\`                                 | e.g. `100` for 100‑px grid; `off` or `0` to disable |

---

## Notes on coordinates & targeting

* **`x/y`** are **normalized** (0..1) coordinates relative to the primary screen.
* **`x_px/y_px`** are **absolute pixels**.
* When clicking inside an **AIM** region, the model must provide explicit coordinates (`x/y` or `x_px/y_px`) **inside** that region. If coordinates are missing or outside AIM, the app ignores the click.
* If the **grid** is enabled, use it to read exact pixel positions.

---

## Examples

### Task (control loop)

```
open Edge, go to Google.com, and search for the term 'życie'
```

### Q\&A (screenshot analysis)

```
/ask where do you see the Edge app icon?
```

---

## Safety, limitations & disclaimer

This is an **experimental** project built for exploration and learning. It simulates real input on your machine and may act unpredictably (mis‑clicks, wrong targets, etc.).

* Use on a **throwaway VM** or a non‑critical environment when possible.
* Review and sandbox tasks.
* You run it **at your own risk**. No warranty of any kind.

---

## License

MIT.
