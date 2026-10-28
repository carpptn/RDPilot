# Contributing to RDPilot

Thank you for helping improve RDPilot. Bug reports, focused pull requests,
documentation improvements, and reproducible desktop-automation scenarios are
all welcome.

## Before you start

- Use Windows 10 or Windows 11 with the
  [.NET 10 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/10.0).
- Read the [safety model](README.md#safety-model) before running the agent.
- Use a disposable VM or non-critical environment for behavioral testing.
- Discuss large behavioral or architectural changes in an issue before investing
  substantial implementation time.

## Development workflow

Fork the repository, create a focused branch, and restore and build the solution:

```powershell
dotnet restore .\RDPilot.sln
dotnet build .\RDPilot.sln -c Release --no-restore
```

Run the deterministic self-test suite without desktop interaction or an API key:

```powershell
dotnet run --project .\RDPilot.SelfTests\RDPilot.SelfTests.csproj -c Release --no-build
```

When changing runtime behavior, add or update a deterministic self-test whenever
practical. Keep changes scoped, preserve nullable-reference-type safety, and
avoid unrelated formatting changes.

## Source layout and test boundaries

Source folders map directly to namespaces (`RDPilot.Control`,
`RDPilot.Observation`, and so on):

| Folder | Responsibility |
| --- | --- |
| `Configuration` | Typed `AppOptions`, profiles, environment and CLI settings |
| `Runtime` | Per-run `RunContext`, counters and the time provider |
| `Control` | Loop coordination, observation/decision/execution/completion stages, policies and batching |
| `Observation` | Screenshot preparation, image comparison and adaptive observation |
| `Desktop` | Native input, screen capture, console and UI Automation |
| `OpenAI` | Request construction, response parsing and HTTP transport |
| `Persistence` | Recovery memory, prompt history and atomic text storage |
| `Diagnostics` | Run logs, request artifacts, replay and analysis |
| `Models` | Shared request/response and domain data types |
| `Properties` | Assembly metadata |

`Program.cs` owns application startup. `ControlLoopService.RunOnce` coordinates
four stages; `ControlRunState` and `ControlStepState` carry state between them.
Run state is initialized by its constructor; step values that require an earlier
stage have guarded accessors. A `ControlDecision` owns its action and follow-up
queue for one step. `BatchPlanner` applies batching and directional policies;
`ControlActionOutcome` accounts for success, failure, and cancellation across
single and grouped actions through `ActionExecutionRecord`, including their own frame and semantic context baselines. `TurnStateStore` owns bounded visual prototypes, navigation positions and edges. `RecoveryStoreFormat` checks versions before payload deserialization; `RecoveryStoreReader` and `RecoveryStoreNormalizer` then validate and migrate supported stores, distinguishing missing, unreadable and unsupported files. Replay version 2 records the live detector's bounded execution history and optional content-free features from `LoopSemanticClassifier`. `TurnVisualAnalysis` owns pure color sampling and pixel change metrics. Follow-ups never live in `RunContext`.
Only values used by more than one stage belong in step state. Helper policies
live in `ControlActionPolicy`, `ControlGoalPolicy`, `ControlLoopGuards`, and
`ControlHistory`. A run snapshots its configuration, restores the prior context
on disposal, and rejects overlapping runs because the console and desktop are
process-wide resources.

Self-tests are grouped into `Behavior`, `Regression`, `Integration`, and
`Support`. Regression tests replace HTTP, native input (including cursor movement and clipboard writes), and screen capture with
in-memory responses, recorded events, and synthetic pixels. `RuntimeClock`
accepts a `TimeProvider`; `AtomicTextStore` exposes file operations for injected
I/O failures. Tests restore substituted dependencies in `finally` blocks.

The default suite makes no API calls and sends no desktop input. A separate,
explicit smoke test uses the existing `OPENAI_API_KEY` and sends a generated
image to the configured default model; this incurs API usage:

```powershell
dotnet run --project .\RDPilot.SelfTests\RDPilot.SelfTests.csproj -c Release -- --live-api-smoke
```

Save source and documentation as UTF-8 with BOM and CRLF line endings.

## Reporting agent failures

Include enough context to reproduce the behavior:

- the goal and selected runtime profile,
- the application and relevant display configuration,
- the smallest useful log excerpt,
- sanitized screenshots or replay artifacts when appropriate,
- the expected and observed result.

Never publish API keys, credentials, private documents, personal information, or
unredacted request and response payloads. Use the private process described in
[SECURITY.md](SECURITY.md) for vulnerabilities.

## Pull requests

Before opening a pull request:

- build the solution in Release configuration,
- run all self-tests,
- update documentation and `appsettings.json` when configuration changes,
- explain user-visible behavior and safety implications,
- keep each pull request focused on one coherent change.

By contributing, you agree that your contribution is licensed under the
repository's [GNU AGPL v3 license](LICENSE.txt).

## Creating a release

Maintainers can test release packaging without publishing from
**Actions → Release → Run workflow**. The dry run validates the build and
self-tests, creates the portable ZIP, and retains it as a workflow artifact for
seven days.

To publish a release, create and push a SemVer-style tag from a verified commit:

```powershell
git tag -a v0.2.0 -m "RDPilot v0.2.0"
git push origin v0.2.0
```

The release workflow verifies the solution, creates the `win-x64` portable ZIP,
checks its required contents, calculates its SHA-256 digest, and publishes a
GitHub Release with generated release notes. Prerelease tags such as
`v0.2.0-beta.1` are automatically marked as prereleases.
