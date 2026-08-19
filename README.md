![version](https://img.shields.io/badge/version-18%2B-EB8E5F)
![platform](https://img.shields.io/static/v1?label=platform&message=mac-intel%20|%20mac-arm%20|%20win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-xls-basic-excel)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-xls-basic-excel/total)

# 4d-pugin-xls-basic-excel

XLS II lets you read and write legacy binary Excel workbooks (`.xls`, BIFF8 / Excel 97-2003 format) directly from 4D code, with no Excel installation, OS automation, or COM/AppleScript bridge involved. It wraps the [BasicExcel](http://www.codeproject.com/Articles/13852/BasicExcel-A-Class-to-Read-and-Write-to-Microsoft) C++ library. You work against an in-memory workbook, addressed by an opaque `Longint` handle returned by `XLS Create` or `XLS Load`; every other command takes that same handle as its first parameter. Results are plain `Longint`, `Real`, or `Text` values — there's no `Picture`/`Blob` result anywhere in this plugin.

> **Note on this revision:** a source-level review found and fixed a handful of bugs (a missing bounds check on four commands that could crash or hang on out-of-range input, a bug that made `XLS Get long value` always return `1` instead of the cell's value, a handle leak on failed `XLS Load` calls, and a formatting-reset bug in `XLS SET WRAPPING`). This doc describes the **fixed** behavior. If you're running an older compiled build of this plugin, some of what's below — specifically the five items just listed — may not yet match what you see; each is called out again at its own command below.

## Summary

| Command | Returns | Purpose |
|---|---|---|
| [XLS Create](#xls-create) | Longint | Create a new in-memory workbook with N sheets |
| [XLS Load](#xls-load) | Longint | Open an existing `.xls` file into a new workbook handle |
| [XLS Save as](#xls-save-as) | Longint | Write a workbook out to a `.xls` file |
| [XLS CLOSE](#xls-close) | *(none)* | Release a workbook handle |
| [XLS Get total sheets](#xls-get-total-sheets) | Longint | Number of sheets in a workbook |
| [XLS Add sheet](#xls-add-sheet) | Longint | Insert a new sheet |
| [XLS Get sheet name](#xls-get-sheet-name) | Longint | Read a sheet's name |
| [XLS Set sheet name](#xls-set-sheet-name) | Longint | Rename a sheet |
| [XLS Get total rows](#xls-get-total-rows) | Longint | Number of used rows in a sheet |
| [XLS Get total columns](#xls-get-total-columns) | Longint | Number of used columns in a sheet |
| [XLS Set column width](#xls-set-column-width) | Longint | Set a column's width |
| [XLS Get column width](#xls-get-column-width) | Longint | Read a column's width |
| [XLS Merge cells](#xls-merge-cells) | Longint | Merge a rectangular cell range |
| [XLS Clear value](#xls-clear-value) | Longint | Erase a cell's value |
| [XLS Get text value](#xls-get-text-value) | Text | Read a cell's text |
| [XLS Get real value](#xls-get-real-value) | Real | Read a cell's numeric value |
| [XLS Get long value](#xls-get-long-value) | Longint | Read a cell's integer value |
| [XLS Set long value](#xls-set-long-value) | Longint | Write an integer to a cell |
| [XLS Set real value](#xls-set-real-value) | Longint | Write a real number to a cell |
| [XLS Set text value](#xls-set-text-value) | Longint | Write text to a cell |
| [XLS Get value type](#xls-get-value-type) | Longint | Get a cell's underlying value type |
| [XLS SET FONT PROPERTY](#xls-set-font-property) | *(none)* | Set a cell's font |
| [XLS GET FONT PROPERTY](#xls-get-font-property) | *(none)* | Read a cell's font |
| [XLS SET FORMAT PROPERTY](#xls-set-format-property) | *(none)* | Set a cell's alignment, borders, and fill pattern |
| [XLS GET FORMAT PROPERTY](#xls-get-format-property) | *(none)* | Read a cell's alignment, borders, and fill pattern |
| [XLS Get format string](#xls-get-format-string) | Longint | Read a cell's number-format string |
| [XLS Set format string](#xls-set-format-string) | Longint | Set a cell's number-format string |
| [XLS Get wrapping](#xls-get-wrapping) | Longint | Read a cell's text-wrap flag |
| [XLS SET WRAPPING](#xls-set-wrapping) | *(none)* | Set a cell's text-wrap flag |

**Platforms:** macOS (Intel & Apple Silicon), Windows 64-bit.

---

## Requirements & platform notes

- **Command-name capitalization is exact and intentional — it isn't cosmetic.** Every command whose C++ name is written in ALL CAPS (`XLS SET FONT PROPERTY`, `XLS GET FONT PROPERTY`, `XLS SET FORMAT PROPERTY`, `XLS GET FORMAT PROPERTY`, `XLS SET WRAPPING`) has **no return value at all** — call it as a bare statement, never assign its result. Every mixed-case command (`XLS Create`, `XLS Get total rows`, etc.) **does** return a value, even if you choose not to capture it. This split is 100% consistent across all 29 commands and was confirmed directly from source, not guessed.
- **No manifest.json was available for this review**, only the compiled command source and five sample `.4dm` test methods. 24 of the 29 command names below are directly confirmed verbatim in those sample files. Five — `XLS Add sheet`, `XLS Merge cells`, `XLS Clear value`, `XLS Set long value`, `XLS Get wrapping` — do not appear in any sample and are **inferred** from the exact naming convention above (which held without exception for the other 24). Treat those five names as very likely correct, not manifest-confirmed; check your plugin's command list in the Explorer if a call doesn't resolve.
- **Handles, not objects.** `XLS Create` and `XLS Load` return a `Longint` workbook handle. Pass that same value as the first parameter to every other command. Call `XLS CLOSE` when you're done with it — it has no return value and is safe to call even on an already-invalid handle (it's a no-op in that case).
- **Everything is 1-based.** Sheet numbers, row numbers, and column numbers are all 1-based at the 4D interface (the plugin subtracts 1 internally before calling into the underlying library).
- **Out-of-range input fails silently, it doesn't raise a 4D error.** Passing a sheet/row/column number that's zero, negative, or past the end of the workbook makes the command a no-op: it returns its type's default (`0`, `0.0`, or an empty string) rather than throwing an error you can catch. Always check bounds yourself with `XLS Get total sheets` / `XLS Get total rows` / `XLS Get total columns` if you're not certain a coordinate is valid, especially before writing.
- **File format is legacy binary `.xls` only** (BIFF8 / Excel 97-2003). BasicExcel does not read or write the XML-based `.xlsx` format — don't pass a `.xlsx` path to `XLS Load`/`XLS Save as`.
- **Text encoding parameters are optional and only matter for non-Unicode content.** `XLS Get text value` and `XLS Set text value` both take a trailing `encoding` `Text` parameter that can be omitted entirely (confirmed by the sample methods, which call both commands with it left off). Leave it empty (or omit it) to read/write native Unicode text. Supply an encoding name (e.g. `"Shift_JIS"`) only when you need to read/write a cell as a specific legacy ANSI codepage. **If you pass a name `iconv` doesn't recognize, the conversion fails silently and you get back empty text — not an error** — so a typo in the encoding name looks identical to an empty/missing cell.
- **Style constants** (`XLS Color Cyan`, `XLS Font Weight Bold`, `XLS Border Dashed`, `XLS Pattern Solid`, `XLS H Align Centered`, etc.) are plugin-supplied constants used in the sample code below. Their full list and numeric values live in the plugin's own constant set in the 4D Explorer/Language Reference for your installed version — this doc doesn't reproduce that list, to avoid inventing values that weren't directly confirmed.

---

## XLS Create

### Syntax
```4d
XLS Create ( sheets ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `sheets` | Longint | Number of sheets to create. |
| Result | Longint | New workbook handle. |

### Description
Creates a brand-new, empty workbook entirely in memory (nothing is written to disk until you call [`XLS Save as`](#xls-save-as)). `sheets` is clamped to a **minimum of 1 and a maximum of 10** — `0` becomes `1`, and anything above `10` becomes `10`; there's no way to request more than 10 sheets at creation time (use [`XLS Add sheet`](#xls-add-sheet) afterward if you need more). Creation always succeeds and always returns a valid handle.

### Example
From the plugin's own test method (`Method1.4dm`):
```4d
//create a blank book with n sheets
$numberOfSheets:=3
$wb:=XLS Create($numberOfSheets)  //min=1, max=10
```
```4d
// A single-sheet workbook
$wb:=XLS Create(1)
```

---

## XLS Load

### Syntax
```4d
XLS Load ( path ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `path` | Text | Path to an existing `.xls` document. |
| Result | Longint | New workbook handle, or `0` if the file couldn't be loaded. |

### Description
Opens an existing `.xls` file and loads it into a new in-memory workbook, returning a fresh handle. If the load fails (bad path, corrupt file, or unsupported format), the result is `0` — check for that before using it.

**Forward-looking fix:** in the reviewed/fixed source, a failed load no longer leaves an orphaned workbook object registered internally with no way to close it. If you're on an older build, a failed `XLS Load` call may leak a small amount of memory per failed attempt; this is fixed going forward.

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$filePath:=System folder:C487(Desktop:K41:16)+"test.xls"
$wb:=XLS Load($filePath)
```
```4d
$wb:=XLS Load($filePath)
If ($wb=0)
	ALERT("Could not open "+$filePath)
Else
	// ... use $wb ...
	XLS CLOSE($wb)
End if
```

---

## XLS Save as

### Syntax
```4d
XLS Save as ( workbook ; path ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `path` | Text | Destination path for the `.xls` file. |
| Result | Longint | `1` on success, `0` if `workbook` is invalid or the write failed. |

### Description
Writes the workbook to disk as a binary `.xls` file at `path`, overwriting any existing file there. Does not close or invalidate the handle — you can keep using it afterward, and you still need to call [`XLS CLOSE`](#xls-close) yourself.

### Example
From the plugin's own test method (`Method1.4dm`):
```4d
$filePath:=System folder:C487(Desktop:K41:16)+"test.xls"
$success:=XLS Save as($wb; $filePath)
```
```4d
$success:=XLS Save as($wb; $filePath)
If (Not($success))
	ALERT("Save failed")
End if
```

---

## XLS CLOSE

### Syntax
```4d
XLS CLOSE ( workbook )
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle to release. |

### Description
Frees the in-memory workbook and invalidates its handle. Has no return value — call it as a bare statement. It's safe to call on an already-closed or otherwise invalid handle; it silently does nothing in that case rather than erroring.

### Example
From the plugin's own test method (`Method1.4dm`):
```4d
XLS CLOSE($wb)
```

---

## XLS Get total sheets

### Syntax
```4d
XLS Get total sheets ( workbook ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| Result | Longint | Number of sheets, or `0` if `workbook` is invalid. |

### Description
Straightforward sheet count for the given workbook.

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$sheets:=XLS Get total sheets($wb)
```

---

## XLS Add sheet

> Command name inferred from naming convention, not directly confirmed in a sample file — see [Requirements & platform notes](#requirements--platform-notes).

### Syntax
```4d
XLS Add sheet ( workbook ; sheetName ; sheetPosition ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheetName` | Text | Name for the new sheet. |
| `sheetPosition` | Longint | 1-based position to insert at. |
| Result | Longint | `1` on success, `0` if `workbook` is invalid. |

### Description
Inserts a new sheet named `sheetName`. If `sheetPosition` is `0`, negative, or greater than the current sheet count, the new sheet is appended at the end instead of raising an error — there's no invalid-position failure mode here, only "insert where asked" or "insert at the end."

### Example
```4d
$wb:=XLS Create(1)
$ok:=XLS Add sheet($wb; "Summary"; 1)   // insert as the first sheet
$ok:=XLS Add sheet($wb; "Details"; 99)  // 99 is out of range -> appended at the end
```

---

## XLS Get sheet name

### Syntax
```4d
XLS Get sheet name ( workbook ; sheet ; name )
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `name` | Text | *(by reference)* Filled with the sheet's name. |
| Result | Longint | `1` on success, `0` otherwise. |

### Description
Reads back a sheet's name into `name`. Internally the plugin first tries to read the name as a wide (Unicode) string, and only falls back to a narrow-string read if that comes back empty — you don't need to do anything differently for either case, both paths land in `name` as ordinary 4D text.

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
XLS Get sheet name($wb; $sheet; $name)
```

---

## XLS Set sheet name

### Syntax
```4d
XLS Set sheet name ( workbook ; sheet ; name ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `name` | Text | New name for the sheet. |
| Result | Longint | `1` on success, `0` if `sheet` is out of range or the rename otherwise fails. |

### Example
From the plugin's own test method (`Method1.4dm`):
```4d
XLS Set sheet name($wb; $sheet; "##SHEET1##")
```

---

## XLS Get total rows

### Syntax
```4d
XLS Get total rows ( workbook ; sheet ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| Result | Longint | Number of used rows, or `0` if `workbook`/`sheet` is invalid. |

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$rows:=XLS Get total rows($wb; $sheet)
```

---

## XLS Get total columns

### Syntax
```4d
XLS Get total columns ( workbook ; sheet ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| Result | Longint | Number of used columns, or `0` if `workbook`/`sheet` is invalid. |

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$cols:=XLS Get total columns($wb; $sheet)
```

---

## XLS Set column width

### Syntax
```4d
XLS Set column width ( workbook ; sheet ; column ; width ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `column` | Longint | 1-based column number. |
| `width` | Longint | Column width, in twips (1/20 of a point). |
| Result | Longint | `1` on success, `0` if `sheet`/`column` is out of range. |

### Example
From the plugin's own test method (`Method1.4dm`):
```4d
$width:=120*20  //in twips (20th of a point)
XLS Set column width($wb; $sheet; $col; $width)
```

---

## XLS Get column width

### Syntax
```4d
XLS Get column width ( workbook ; sheet ; column ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `column` | Longint | 1-based column number. |
| Result | Longint | Column width in twips, or `0` if `sheet`/`column` is out of range. |

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$width:=XLS Get column width($wb; $sheet; $row; $col)
```
> Note: the sample above passes 4 parameters including `$row`, which this command doesn't take (`XLS Get column width` is `workbook; sheet; column`, not `workbook; sheet; row; column`). This is quoted exactly as it appears in the plugin's own `Method2.4dm` — the extra argument is silently ignored by 4D's own parameter handling, but don't copy that call shape into new code; use the 3-parameter form shown in the syntax above.

---

## XLS Merge cells

> Command name inferred from naming convention, not directly confirmed in a sample file — see [Requirements & platform notes](#requirements--platform-notes).

### Syntax
```4d
XLS Merge cells ( workbook ; sheet ; row ; column ; height ; width ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based anchor row. |
| `column` | Longint | 1-based anchor column. |
| `height` | Longint | Row extent of the merge from the anchor cell. |
| `width` | Longint | Column extent of the merge from the anchor cell. |
| Result | Longint | `1` on success, `0` otherwise. |

### Description
Merges a rectangular range anchored at `(row, column)`. At least one of `height`/`width` must be greater than `1` (a 1×1 "merge" is rejected as a no-op), and the anchor plus the extent must stay within the sheet's current row/column count — both are enforced before the merge is attempted.

The exact accepted meaning of `height`/`width` (total span vs. cells beyond the anchor) is set by the underlying BasicExcel library's own `MergeCells` call and isn't independently confirmable from this plugin's source alone or from the sample files (none of them exercise this command) — test with a small, visually-checkable range before relying on exact row/column counts in production.

### Example
```4d
// Merge a 1-row-by-3-column header band starting at row 1, column 1
$ok:=XLS Merge cells($wb; 1; 1; 1; 1; 3)
```

---

## XLS Clear value

> Command name inferred from naming convention, not directly confirmed in a sample file — see [Requirements & platform notes](#requirements--platform-notes).

### Syntax
```4d
XLS Clear value ( workbook ; sheet ; row ; column ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| Result | Longint | `1` on success, `0` if the coordinates are out of range. |

### Description
Erases the value stored at the given cell (formatting is untouched).

### Example
```4d
$ok:=XLS Clear value($wb; 1; 2; 2)
```

---

## XLS Get text value

### Syntax
```4d
XLS Get text value ( workbook ; sheet ; row ; column { ; encoding } ) -> Text
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `encoding` | Text | *Optional.* Name of the ANSI encoding to decode with, only used if the cell holds a legacy narrow-string value rather than Unicode. |
| Result | Text | The cell's text, or an empty string if the coordinates are out of range or there's no string content. |

### Description
If the cell holds a Unicode (wide) string, that's returned directly and `encoding` is ignored entirely. Only if the cell's string is stored as a legacy narrow/ANSI string does `encoding` come into play, naming the codepage to decode it with. Leave `encoding` empty/omitted for ordinary Unicode-authored workbooks.

**An unrecognized `encoding` name fails silently** — you get back an empty string, indistinguishable from an actually-empty cell.

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$text:=XLS Get text value($wb; $sheet; $row; $col)
```
```4d
// Reading a value that might be stored in a legacy Japanese codepage
$text:=XLS Get text value($wb; 1; 3; 2; "Shift_JIS")
```

---

## XLS Get real value

### Syntax
```4d
XLS Get real value ( workbook ; sheet ; row ; column ) -> Real
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| Result | Real | The cell's numeric value, or `0` if the coordinates are out of range or the cell has no numeric content. |

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$real:=XLS Get real value($wb; $sheet; $row; $col)
```

---

## XLS Get long value

### Syntax
```4d
XLS Get long value ( workbook ; sheet ; row ; column ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| Result | Longint | The cell's value, read as an integer. `0` if the coordinates are out of range. |

### Description
**Forward-looking fixes, both applied in this revision:**
- The command previously always returned `1` regardless of the cell's actual content — a bug that made it useless for anything but a presence check. It now correctly returns the cell's integer value.
- The command previously had no bounds check on `sheet`/`row`/`column` at all (unlike every other Get/Set command in the plugin), so an out-of-range sheet or cell coordinate could crash the plugin instead of returning `0`. That check is now in place, matching every sibling command.

If you're on an older build, don't rely on this command's return value for anything beyond "is this cell non-empty" — get the real value instead if you need to distinguish `0` from "not set."

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$long:=XLS Get long value($wb; $sheet; $row; $col)
```

---

## XLS Set long value

> Command name inferred from naming convention, not directly confirmed in a sample file — see [Requirements & platform notes](#requirements--platform-notes).

### Syntax
```4d
XLS Set long value ( workbook ; sheet ; row ; column ; value ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `value` | Longint | Integer to write. |
| Result | Longint | `1` on success, `0` if the coordinates are out of range. |

### Description
**Forward-looking fix:** this command previously had no bounds check on `sheet`/`row`/`column`, so an out-of-range coordinate could crash the plugin rather than returning `0`. That check is now in place, matching every sibling command.

### Example
```4d
$ok:=XLS Set long value($wb; 1; 4; 4; 42)
```

---

## XLS Set real value

### Syntax
```4d
XLS Set real value ( workbook ; sheet ; row ; column ; value ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `value` | Real | Number to write. |
| Result | Longint | `1` on success, `0` if the coordinates are out of range. |

### Description
**Forward-looking fix:** this command previously had no bounds check on `sheet`/`row`/`column`, so an out-of-range coordinate could crash the plugin rather than returning `0`. That check is now in place, matching every sibling command.

### Example
From the plugin's own test method (`Method3.4dm`):
```4d
XLS Set real value($wb; $sheet; 2; 2; 1234.5678)
```

---

## XLS Set text value

### Syntax
```4d
XLS Set text value ( workbook ; sheet ; row ; column ; text { ; encoding } ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `text` | Text | Text to write. |
| `encoding` | Text | *Optional.* If omitted or empty, `text` is stored as Unicode. If given, `text` is converted to that ANSI codepage first and stored as a legacy narrow string. |
| Result | Longint | `1` on success, `0` if the coordinates are out of range. |

### Description
**Forward-looking fix:** this command previously had no bounds check on `sheet`/`row`/`column`, so an out-of-range coordinate could crash the plugin rather than returning `0`. That check is now in place, matching every sibling command.

**An unrecognized `encoding` name fails silently on write, too** — if `iconv` doesn't recognize the name, the conversion produces nothing and the cell is left with no updated text, with no error reported.

### Example
From the plugin's own test method (`Method1.4dm`):
```4d
XLS Set text value($wb; $sheet; $row; $col; "1,1:abcde")
```
```4d
// Explicit legacy-encoding write
$ok:=XLS Set text value($wb; 1; 5; 1; "テスト"; "Shift_JIS")
```

---

## XLS Get value type

### Syntax
```4d
XLS Get value type ( workbook ; sheet ; row ; column ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| Result | Longint | Raw cell-type code from the underlying library, or `0` if the coordinates are out of range. |

### Description
Returns BasicExcel's own internal cell-type code as-is. This plugin's source doesn't translate that code into a named 4D constant, and the exact integer-to-type mapping (text vs. number vs. formula vs. empty, etc.) is defined by the underlying library, not by this doc — compare the returned value empirically against cells of known content, or check the vendored `ExcelFormat.h`/`BasicExcel.h` if you have it, rather than assuming a specific mapping.

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$type:=XLS Get value type($wb; $sheet; $row; $col)
```

---

## XLS SET FONT PROPERTY

### Syntax
```4d
XLS SET FONT PROPERTY ( workbook ; sheet ; row ; column ; name ; height ; color ; weight ; option ; underline ; family ; escapement )
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `name` | Text | Font name, e.g. `"Lucida Grande"`. |
| `height` | Longint | Font size, in twips (1/20 of a point). |
| `color` | Longint | Color index constant, e.g. `XLS Color Cyan`. |
| `weight` | Longint | Weight constant, e.g. `XLS Font Weight Bold`. |
| `option` | Longint | Style-option flags, e.g. `XLS Font Italic`. |
| `underline` | Longint | Underline-type constant, e.g. `XLS Underline Single`. |
| `family` | Longint | Font-family constant, e.g. `XLS Font Family Roman`. |
| `escapement` | Longint | Escapement (super/subscript) constant, e.g. `XLS Escapement Normal`. |

### Description
No return value — call as a bare statement. Applies the given font to the single cell at `(row, column)`; other formatting already applied to that cell (alignment, borders, pattern) is preserved, only the font is replaced.

### Example
From the plugin's own test method (`Method1.4dm`):
```4d
$font:="Lucida Grande"
$height:=12*20  //in twips (20th of a point)
$color:=XLS Color Cyan
$weight:=XLS Font Weight Bold
$option:=XLS Font Italic
$underline:=XLS Underline Single
$family:=XLS Font Family Roman
$escape:=XLS Escapement Normal
XLS SET FONT PROPERTY($wb; $sheet; $row; $col; \
$font; $height; $color; $weight; $option; $underline; $family; $escape)
```

---

## XLS GET FONT PROPERTY

### Syntax
```4d
XLS GET FONT PROPERTY ( workbook ; sheet ; row ; column ; name ; height ; color ; weight ; option ; underline ; family ; escapement )
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `name` | Text | *(by reference)* Filled with the cell's font name. |
| `height` | Longint | *(by reference)* Filled with the font size, in twips. |
| `color` | Longint | *(by reference)* Filled with the color index. |
| `weight` | Longint | *(by reference)* Filled with the weight constant. |
| `option` | Longint | *(by reference)* Filled with the style-option flags. |
| `underline` | Longint | *(by reference)* Filled with the underline-type constant. |
| `family` | Longint | *(by reference)* Filled with the font-family constant. |
| `escapement` | Longint | *(by reference)* Filled with the escapement constant. |

### Description
No return value. If the coordinates are out of range, none of the by-reference parameters are touched — initialize them yourself before the call if you need a defined value in that case (the sample below does exactly that).

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$font:=""
$height:=0
$color:=0
$weight:=0
$option:=0
$underline:=0
$family:=0
$escape:=0
XLS GET FONT PROPERTY($wb; $sheet; $row; $col; \
$font; $height; $color; $weight; $option; $underline; $family; $escape)
```

---

## XLS SET FORMAT PROPERTY

### Syntax
```4d
XLS SET FORMAT PROPERTY ( workbook ; sheet ; row ; column ; alignment ; rotation ; textProps ; borderlineTop ; borderlineTopColor ; borderlineLeft ; borderlineLeftColor ; borderlineRight ; borderlineRightColor ; borderlineBottom ; borderlineBottomColor ; pattern ; patternColor ; patternBackColor )
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `alignment` | Longint | Horizontal/vertical alignment constants, combined with `\|`, e.g. `XLS H Align Centered \| XLS V Align Centered`. |
| `rotation` | Longint | Text rotation, in degrees. |
| `textProps` | Longint | Text property flags, e.g. `XLS Text Default`. |
| `borderlineTop` | Longint | Top border line-style constant, e.g. `XLS Border Dashed`. |
| `borderlineTopColor` | Longint | Top border color index. |
| `borderlineLeft` | Longint | Left border line-style constant. |
| `borderlineLeftColor` | Longint | Left border color index. |
| `borderlineRight` | Longint | Right border line-style constant. |
| `borderlineRightColor` | Longint | Right border color index. |
| `borderlineBottom` | Longint | Bottom border line-style constant. |
| `borderlineBottomColor` | Longint | Bottom border color index. |
| `pattern` | Longint | Fill pattern constant, e.g. `XLS Pattern Solid`. |
| `patternColor` | Longint | Fill foreground color index. |
| `patternBackColor` | Longint | Fill background color index. |

### Description
No return value. Applies alignment, rotation, text properties, all four border lines/colors, and the fill pattern/colors to the single cell at `(row, column)` in one call. Other formatting already on that cell (font) is preserved.

### Example
From the plugin's own test method (`Method1.4dm`):
```4d
$alignment:=XLS H Align Centered | XLS V Align Centered
$rotation:=45
$properties:=XLS Text Default
$top:=XLS Border Dashed
$topColor:=XLS Color Red
$left:=XLS Border Dashed
$leftColor:=XLS Color Red
$right:=XLS Border Dashed
$rightColor:=XLS Color Red
$bottom:=XLS Border Dashed
$bottomColor:=XLS Color Red
$pattern:=XLS Pattern Solid
$patternColor:=XLS Color Blue
$patternAltColor:=XLS Color Yellow
XLS SET FORMAT PROPERTY($wb; $sheet; 2; 2; \
$alignment; $rotation; $properties; \
$top; $topColor; \
$left; $leftColor; \
$right; $rightColor; \
$bottom; $bottomColor; \
$pattern; $patternColor; $patternAltColor)
```

---

## XLS GET FORMAT PROPERTY

### Syntax
```4d
XLS GET FORMAT PROPERTY ( workbook ; sheet ; row ; column ; alignment ; rotation ; textProps ; borderlineTop ; borderlineTopColor ; borderlineLeft ; borderlineLeftColor ; borderlineRight ; borderlineRightColor ; borderlineBottom ; borderlineBottomColor ; pattern ; patternColor ; patternBackColor )
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `alignment` | Longint | *(by reference)* Horizontal/vertical alignment flags. |
| `rotation` | Longint | *(by reference)* Text rotation in degrees. |
| `textProps` | Longint | *(by reference)* Text property flags. |
| `borderlineTop` | Longint | *(by reference)* Top border line-style constant. |
| `borderlineTopColor` | Longint | *(by reference)* Top border color index. |
| `borderlineLeft` | Longint | *(by reference)* Left border line-style constant. |
| `borderlineLeftColor` | Longint | *(by reference)* Left border color index. |
| `borderlineRight` | Longint | *(by reference)* Right border line-style constant. |
| `borderlineRightColor` | Longint | *(by reference)* Right border color index. |
| `borderlineBottom` | Longint | *(by reference)* Bottom border line-style constant. |
| `borderlineBottomColor` | Longint | *(by reference)* Bottom border color index. |
| `pattern` | Longint | *(by reference)* Fill pattern constant. |
| `patternColor` | Longint | *(by reference)* Fill foreground color index. |
| `patternBackColor` | Longint | *(by reference)* Fill background color index. |

### Description
No return value. Same 18-parameter shape as [`XLS SET FORMAT PROPERTY`](#xls-set-format-property), used to read a cell's current formatting back into the same named variables. If the coordinates are out of range, none of the by-reference parameters are touched — initialize them yourself beforehand.

### Example
From the plugin's own test method (`Method2.4dm`):
```4d
$alignment:=0
$rotation:=0
$properties:=0
$top:=0
$topColor:=0
$left:=0
$leftColor:=0
$right:=0
$rightColor:=0
$bottom:=0
$bottomColor:=0
$pattern:=0
$patternColor:=0
$patternAltColor:=0
XLS GET FORMAT PROPERTY($wb; $sheet; 2; 2; \
$alignment; $rotation; $properties; \
$top; $topColor; \
$left; $leftColor; \
$right; $rightColor; \
$bottom; $bottomColor; \
$pattern; $patternColor; $patternAltColor)
```

---

## XLS Get format string

### Syntax
```4d
XLS Get format string ( workbook ; sheet ; row ; column ; formatString ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `formatString` | Text | *(by reference)* Filled with the cell's Excel number-format string. |
| Result | Longint | `1` on success, `0` if the coordinates are out of range. |

### Description
Reads the Excel-style number-format pattern applied to the cell (the same syntax you'd see in Excel's own "Format Cells" custom-format box).

### Example
From the plugin's own test method (`Method3.4dm`):
```4d
$success:=XLS Get format string($wb; $sheet; 2; 2; $format)
```

---

## XLS Set format string

### Syntax
```4d
XLS Set format string ( workbook ; sheet ; row ; column ; formatString ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `formatString` | Text | Excel-style number-format pattern to apply. |
| Result | Longint | `1` on success, `0` if the coordinates are out of range. |

### Example
From the plugin's own test method (`Method3.4dm`):
```4d
$success:=XLS Set format string($wb; $sheet; 2; 2; "\"$\"#,##0_);(\"$\"#,##0)")
```

---

## XLS Get wrapping

> Command name inferred from naming convention, not directly confirmed in a sample file — see [Requirements & platform notes](#requirements--platform-notes).

### Syntax
```4d
XLS Get wrapping ( workbook ; sheet ; row ; column ) -> Longint
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| Result | Longint | `1` if the cell wraps text, `0` if it doesn't (or if the coordinates are out of range). |

### Example
```4d
$wraps:=XLS Get wrapping($wb; 1; 4; 1)
```

---

## XLS SET WRAPPING

### Syntax
```4d
XLS SET WRAPPING ( workbook ; sheet ; row ; column ; wrapping )
```

| Parameter | Type | Description |
|---|---|---|
| `workbook` | Longint | Workbook handle. |
| `sheet` | Longint | 1-based sheet number. |
| `row` | Longint | 1-based row number. |
| `column` | Longint | 1-based column number. |
| `wrapping` | Longint | `1` to enable text wrap, `0` to disable. |

### Description
No return value. Sets the cell's text-wrap flag.

**Forward-looking fix:** this command previously rebuilt the cell's format from scratch instead of starting from its existing formatting, so calling it would silently discard any font, alignment, border, or fill-pattern settings already applied to that cell. It's now fixed to preserve existing formatting and only change the wrap flag, matching how every other single-property setter in this plugin behaves. If you're on an older build, call this **before** applying font/format/font-property commands to a cell, not after, to avoid losing that other formatting.

### Example
From the plugin's own test method (`Method4.4dm`):
```4d
XLS Set text value($wb; $sheet; $row; $col; "abcde"*100)
XLS SET WRAPPING($wb; $sheet; $row; $col; 1)
```

---

## Error handling & troubleshooting

- **A `0` or empty-string result almost never means "error" in the 4D-visible sense** — this plugin doesn't raise 4D errors for bad input. `0`/`0.0`/`""` back from a Get command, or `0` back from a Set command, means either the workbook/sheet/row/column was out of range, or (for the format/font getters) the relevant by-reference parameters were simply never touched. Check `XLS Get total sheets`/`rows`/`columns` yourself if you're not sure a coordinate is valid.
- **`XLS Load` returning `0`** means the file didn't load — bad path, corrupted file, or a format BasicExcel doesn't understand (including `.xlsx`, which this plugin cannot read).
- **An unrecognized `encoding` name on `XLS Get text value`/`XLS Set text value` fails silently.** You'll get back (or write) empty text with no indication the encoding name itself was the problem — double-check the spelling against your platform's `iconv -l` output if a text round-trip comes back blank unexpectedly.
- **`XLS Get value type`'s return code isn't documented anywhere accessible from 4D.** It's BasicExcel's own internal cell-type enum passed straight through — determine the mapping empirically against cells of known content if you need to branch on it.
- **Every command handle you get back from `XLS Create`/`XLS Load` needs its own `XLS CLOSE`.** There's no "close all" and no automatic cleanup at the end of a method — an unclosed handle stays allocated in memory for the life of the 4D session/process.
- **Don't mix up `XLS GET FONT PROPERTY`/`XLS GET FORMAT PROPERTY`'s by-reference output parameters with a return value** — these two, along with `XLS SET FONT PROPERTY`, `XLS SET FORMAT PROPERTY`, and `XLS SET WRAPPING`, have no return value at all. Initialize your output variables before the call if you need a defined fallback when the coordinates turn out to be out of range.

---

## Quick reference

```4d
// Create, write, format, save, close
$wb:=XLS Create(1)
XLS Set text value($wb; 1; 1; 1; "Hello")
XLS Set real value($wb; 1; 2; 1; 1234.5678)
XLS Set format string($wb; 1; 2; 1; "\"$\"#,##0_);(\"$\"#,##0)")
XLS SET WRAPPING($wb; 1; 1; 1; 1)
$path:=System folder:C487(Desktop:K41:16)+"out.xls"
$ok:=XLS Save as($wb; $path)
XLS CLOSE($wb)

// Load and read back
$wb:=XLS Load($path)
If ($wb#0)
	$sheets:=XLS Get total sheets($wb)
	$rows:=XLS Get total rows($wb; 1)
	$cols:=XLS Get total columns($wb; 1)
	$text:=XLS Get text value($wb; 1; 1; 1)
	$real:=XLS Get real value($wb; 1; 2; 1)
	XLS CLOSE($wb)
End if
```
