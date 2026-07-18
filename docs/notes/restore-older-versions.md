# Restore Older Versions

The project keeps Git tags and archive branches so older states can be reviewed or restored.

## GitHub URLs

| Resource | URL |
| --- | --- |
| All tags | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tags> |
| All branches | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/branches> |
| Current source history | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/main/docs/source-history> |

## Important Restore Points

| Version | Tag URL | Archive Branch |
| --- | --- | --- |
| `1.1.6` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.6> | `archive/v1.1.6` |
| `1.1.5` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.5> | `archive/v1.1.5` |
| `1.1.4` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.4> | `archive/v1.1.4` |
| `1.1.3` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.3> | `archive/v1.1.3` |
| `1.1.2` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.2> | `archive/v1.1.2` |
| `1.1.1` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.1> | `archive/v1.1.1` |
| `1.1.0` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.1.0> | `archive/v1.1.0` |
| `1.0.9` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.0.9> | `archive/v1.0.9` |
| `1.0.8` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.0.8> | `archive/v1.0.8` |
| `1.0.7` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.0.7> | `archive/v1.0.7` |
| `1.0.6` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.0.6> | `archive/v1.0.6` |
| `1.0.5` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.0.5> | `archive/v1.0.5` |
| `1.0.4` | <https://github.com/rice2k/Macys-Asset-Protection-HEX-Converter-Tool/tree/v1.0.4> | `archive/v1.0.4` |

## View An Older Version On GitHub

1. Open the tag URL.
2. Review the files from that version.
3. Download a ZIP from GitHub if needed.

## Restore Locally With Git

```powershell
git fetch --all --tags
git checkout v1.1.1
```

To return to the current main branch:

```powershell
git checkout main
git pull origin main
```

## Safer Test Branch Restore

Use this when you want to inspect an older version without changing `main`:

```powershell
git fetch --all --tags
git checkout -b test-restore-v1.0.9 v1.0.9
```

## Original HTML Versions

The original browser-based versions are stored under:

[docs/source-history/html-versions](../source-history/html-versions)

Those files are historical reference material. The maintained app is the Windows desktop app on `main`.
