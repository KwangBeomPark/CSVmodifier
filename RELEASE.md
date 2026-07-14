# Release checklist

This project publishes a single Windows executable from `csv_modifier.spec`.

## Required executable naming rule

Every release executable **must begin with `App04_`**.

The PyInstaller spec derives the release asset name from `__version__` in
`csv_modifier.py` and enforces this format:

```text
App04_csv_modifier_v<version>.exe
```

For example, version `1.5.0` must be released as:

```text
App04_csv_modifier_v1.5.0.exe
```

Do not upload an executable that does not follow this prefix rule.

## Release steps

1. Update `__version__` in `csv_modifier.py`.
2. Update the version and executable name in both README files.
3. Run `python -m unittest discover -s tests -v`.
4. Build with `python -m PyInstaller --clean --noconfirm csv_modifier.spec`.
5. Verify `dist/App04_csv_modifier_v<version>.exe` exists and has the app icon.
6. Create Git tag `v<version>` and upload that executable to the matching GitHub release.
