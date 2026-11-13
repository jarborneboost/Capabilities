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
