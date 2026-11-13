# CLAUDE.md

## GitHub Actions

After adding a workflow or changing a `uses:` line, run `gh actions-lock` and commit `.github/workflows/actions.lock` with it; never edit the lockfile by hand. CI's `gh actions-lock --verify` fails when they disagree. Install the extension at the version `ci.yml` pins (`gh extension install github/gh-actions-lock --pin v0.4.1`), and bump both together.
