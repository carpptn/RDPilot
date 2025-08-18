# RDPilot — AI‑Controlled Desktop Agent (Experimental)

**RDPilot** is an experimental, vibe‑coded console app (C# / .NET 9, Windows) that lets a Large Language Model (LLM) operate your desktop by looking at screenshots and emitting actions (keyboard, mouse, scroll, etc.).

* Best results so far with **`gpt-5`**; older models can be **faster**, but are usually less reliable.
* Designed for **Windows 10/11 (x64)**. .NET **9** is required.

> ⚠️ By default the **mouse is disabled** because current models still struggle with precise targeting of on‑screen elements. Keyboard‑first strategies are encouraged. You can enable the mouse with a flag or env var (see below).

> ⚠️ Currently RDPilot captures and operates only the primary display.

---

## Requirements

* **Windows 10/11 (x64)**
* **.NET 9** runtime or SDK (**recommended**)
* **.NET 9.0 Desktop Runtime (Windows Desktop Runtime)**
* OpenAI API KEY

---

## What can it do?

Give the model a high‑level goal and it will iteratively act on your desktop. For example:

```
open Edge, go to Google.com, and search for the term 'life'
```

It also supports “Q\&A on screenshot” via `/ask`, e.g.:

```
/ask where do you see the Edge app icon?
```
---

## How it works (high level)

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
- `keystype_text`  
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
The application executes the given action via **WinAPI (SendInput)**.  

**5. Iterative Loop**  
After execution, a new screenshot is generated and sent back to the model.  
The model decides the next action.  
This process repeats in a loop until the model returns the `done` action, which signals that the initial task goal (from the prompt) has been achieved.  

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

## Running

You can pass a **single goal** as an argument, or start with no args to use **interactive mode**.

```bash
# interactive mode
RDPilot.exe
# then type goals or Q&A like:
#   open Edge, go to Google.com, and search for the term 'life'
#   /ask where do you see the Edge app icon?

# one‑shot goal
RDPilot.exe "open Edge, go to Google.com, and search for the term 'life'"
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

## Notes on coordinates & targeting (mouse)

* **`x/y`** are **normalized** (0..1) coordinates relative to the primary screen.
* **`x_px/y_px`** are **absolute pixels**.
* When clicking inside an **AIM** region, the model must provide explicit coordinates (`x/y` or `x_px/y_px`) **inside** that region. If coordinates are missing or outside AIM, the app ignores the click.
* If the **grid** is enabled, use it to read exact pixel positions.

---

## Examples

### Task (control loop)

```
open Edge, go to Google.com, and search for the term 'life'
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

If you improve this code, I’ll be happy to accept the changes in a PR :)

