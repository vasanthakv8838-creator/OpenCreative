# Contributing to OpenCreative

Thank you for helping make OpenCreative better! Contributions of all sizes are welcome: bug reports, documentation, accessibility improvements, design feedback, and code.

## Before you start

- Read the [README](README.md) to understand the current MVP and its limitations.
- Check existing [issues](../../issues) and [pull requests](../../pulls) to avoid duplicate work.
- For a large change, open an issue first so we can discuss the approach.

## Report a bug or suggest an idea

Use the repository's issue templates. Include clear steps to reproduce bugs, your browser/device, what you expected, and what happened. Do not attach private files or personal information.

## Make a code contribution

1. Fork this repository.
2. Create a focused branch from `main`, for example `fix/vector-export`.
3. Make the smallest complete change that solves the issue.
4. Keep the app local-first. Do not upload users' files or add telemetry without an explicit, documented decision.
5. Use original code and assets, and preserve third-party license notices.
6. Test the affected feature in a modern browser. For microphone or other secure-context APIs, test through HTTPS or localhost.
7. Check mobile behavior where practical; browser APIs and codecs vary by device.
8. Commit your changes with a clear message.
9. Open a pull request against `main`. Explain the problem, the solution, how you tested it, and link related issues.

## Pull request expectations

- Keep each PR focused and explain user-visible changes.
- Update documentation when behavior or setup changes.
- Do not include generated bundles, personal files, secrets, or unrelated formatting changes.
- Be respectful in review discussions and be open to constructive feedback.

## Tests

This MVP currently documents manual browser testing rather than a configured automated test suite. Describe the exact manual checks you performed. If you add tests or tooling, document how to run them.

## Questions

If you are unsure where to begin, open an issue or start a conversation in GitHub Discussions (when enabled). A small question is always welcome.

By participating, you agree to follow our [Code of Conduct](CODE_OF_CONDUCT.md).
