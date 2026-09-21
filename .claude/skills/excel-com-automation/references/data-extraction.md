# Reading data out of Excel efficiently and correctly

## Bulk reads, chunked

Never loop over `ws.Cells(r, c).Value` - each access is a full COM round-trip and is orders of magnitude slower
than a single bulk read. Read a whole rectangle at once:

```python
def read_block(ws, row, col, nrows, ncols):
    v = ws.Range(ws.Cells(row, col), ws.Cells(row + nrows - 1, col + ncols - 1)).Value2
    return [(v,)] if not isinstance(v, tuple) else list(v)   # a 1x1 range returns a scalar, not a tuple
```

For sheets with hundreds of thousands of rows, chunk by row ranges rather than reading the whole sheet in one
call - this keeps any single COM call fast and keeps peak memory bounded. Size the chunk so `rows × cols` stays
in the low millions of cells; this project used a floor of 500 rows and a ceiling chosen so a very wide sheet
still got a reasonable number of rows per chunk:

```python
step = max(500, min(15000, MAX_CELLS_PER_CHUNK // max(ncols, 1)))
```

This comfortably handled a real 246,000-row sheet and a synthetic sheet filled to Excel's hard limit of
1,048,575 data rows.

## `Value2`, not `.Value`

`.Value` applies COM automation's currency/date auto-formatting on the way out, which is slower and can silently
coerce types you didn't ask for. `Value2` gives you Excel's raw representation - plain floats for numbers and
dates (as day-count serials you convert yourself), plain strings, `None` for empty, and specific negative
integers for error cells. See `gotchas.md` #7 for handling error cells and the two date systems.

## Anchor reads at column A, not at the detected "used" start column

If you separately determine the real last row/column (see `gotchas.md` #1), read from column 1 regardless of
where `UsedRange` claims data starts - otherwise your column indices from a header-detection pass and your
column indices from a bulk-data pass can disagree by an offset, which is a subtle, easy-to-miss bug.

## Detecting which file is which, and which columns are which, by content

For a pipeline that ingests files from someone else's export process (not files you control the layout of):

- **Identify the file's role by its header content**, not its filename - filenames often carry a timestamp or
  export-tool suffix that changes run to run, while the *meaning* of "this is the file with columns X, Y, Z"
  doesn't.
- **Normalise headers before matching**: Unicode-normalise (NFKC), strip invisible characters (zero-width
  space, NBSP, soft hyphen, BOM), collapse whitespace/newlines, lower-case. Real exports vary in exactly these
  ways release to release without the underlying meaning changing.
- **Scan a bounded number of header rows** (not just row 1) - some exports have title rows, or two-row grouped
  headers, above the real header line.
- **When nothing matches well enough**, report the single *closest* candidate and name exactly which expected
  column(s) are missing, rather than a bare "this isn't the right file." This is the difference between a
  non-technical user fixing their own export and a support ticket.
- **Never guess ambiguous data** (see `gotchas.md` #9 for dates) - an incorrect-but-plausible guess is worse
  than a clearly reported "couldn't read this" that a human can act on.

## Model/ID keys: normalise before joining

When joining rows across files by some kind of code/ID column, apply the same normalisation as headers (case,
whitespace, invisible characters) *and* handle the case where Excel stored what looks like a text code as a
number (e.g. a numeric-looking model code becomes `12345.0` as a float) - convert integral floats back to plain
integer text before using them as a join key, or joins will silently fail for exactly those rows.
