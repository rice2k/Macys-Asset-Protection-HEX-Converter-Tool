# Input Examples

## Good HEX Input

```text
88984717
88984765
88984130
88981234
```

## Cleaned Spreadsheet Input

These can be cleaned before conversion:

```text
88984765.0
8898-4130
8898 1234
Candidate Name, 88984717
```

Possible Notes / Details:

| Input | Result |
| --- | --- |
| `88984765.0` | Cleaned Excel numeric ID. |
| `8898-4130` | Joined split ID. |
| `Candidate Name, 88984717` | Extracted ID from full text. |
| Duplicate value | Duplicate of an earlier line. |

## Bad Input

```text
88984
8898471700
88984XYZ
No ID Here
```

Common invalid reasons:

- Too short.
- Too long without a clean 8-character token.
- Invalid HEX character.
- Empty line.
- Full text with no usable ID.

## FC/CN To Hex Input

Single reverse lookup accepts:

```text
Facility Code: 34968
Card Number: 18199
```

Unconvert Batch accepts examples like:

```text
34968,18199
34968 18277
FC 34968 CN 18199
```

## Example Conversion

| HEX | Facility Code | Card Number |
| --- | ---: | ---: |
| `88984717` | `34968` | `18199` |
| `88984765` | `34968` | `18277` |
| `88984130` | `34968` | `16688` |

## Validation Colors

| Color | Meaning |
| --- | --- |
| Green | Clean valid row. |
| Yellow | Valid row with a warning, cleanup note, duplicate, or unusual value. |
| Red | Invalid row that needs review. |
