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
