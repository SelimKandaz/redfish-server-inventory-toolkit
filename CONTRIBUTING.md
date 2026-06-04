# Contributing

Thanks for your interest in improving this project.

## Good contributions

- Bug fixes
- Safer defaults
- Better documentation
- Better examples
- Tests
- Clearer setup instructions
- More flexible configuration options

## Before opening a pull request

1. Keep examples sanitized and generic.
2. Do not include private business files or generated local outputs.
3. Keep changes focused.
4. Add or update documentation when behavior changes.
5. Run the available checks before submitting.

## Local checks

For Python projects:

```bash
python -m compileall .
```

For shell scripts:

```bash
bash -n scripts/*.sh
```

For .NET projects:

```bash
dotnet build
```
