# Capabilities

## Usage

Validate the plugin configuration:

```console
claude plugin validate .
```

### Plugins

Add a marketplace:

```console
/plugin marketplace add <HOST>/<YOUR-NAMESPACE>/<YOUR-PROJECT>.git

# For this one:
/plugin marketplace add https://github.com/jarborneboost/capabilities.git
```

Install a plugin:

```console
/plugin install <plugin-name>@<marketplace-name> --scope user    # available to you across projects
/plugin install <plugin-name>@<marketplace-name> --scope project # available from current repo's .claude/settings.json
/plugin install <plugin-name>@<marketplace-name> --scope local   # available only to you, in current repo

# For this one:
/plugin install my-plugin-bundle --marketplace https://github.com/jarborneboost/capabilities.git
```

## Publishing

Publishing means pushing to the default branch (`main`). Users that have added this plugin (and have auto-sync on) will automatically start using updated skills.

1. Bump `version` in the plugin's `plugin.json` (installed plugins only update when the version changes).
2. The version in `package.json` does not follow the version in `plugin.json` (we only use that for internal tooling)
3. Commit and push to `main`.

## Tests

Scripts bundled in skills are tested in `tests/`, outside `plugins/` so the tests are not shipped with the plugins.

Requires [uv](https://docs.astral.sh/uv/).

```console
pnpm test
```

## Docs

- [Agent Skills (overview)](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview)
- [Skill authoring best practices](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/best-practices)
