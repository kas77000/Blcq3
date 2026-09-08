# Reading numbers off images

The numbers arrive as photos. Transcription is the weakest link in the whole
chain, so it gets the most process.

## Convert first

```bash
python .claude/skills/tca-review/scripts/intake.py \
  --in reviews/<name>/images --out reviews/<name>/images/png --tile
```

Handles `.heic`, `.heif`, `.jpg`, `.png`. It applies EXIF rotation, upscales, and
with `--tile` also writes quadrant crops of dense images - a 40-row table is far
more reliable read one quadrant at a time than whole. It writes `inventory.csv`.

If a table is still not legible after tiling, **say so and ask for a better
image**. A re-shot photo costs the user a minute. A misread digit in front of a
client costs a great deal more.

## Transcribe

Read the full image first to see the table's shape, then read the tiles to get
the digits.

Rules:

1. **Transcribe only what is visible.** Unreadable cell → `NA`, and a line in
   `sources.csv` saying which cell and why.
2. **Never infer a digit from the pattern.** A column that reads 1.2, 1.4, 1.?,
   1.8 does not make the missing one 1.6.
3. **Copy the sign exactly**, including a leading minus that a photo may clip. If
   a minus is ambiguous, that is an `NA`.
4. **Record units as printed** - millions, thousands, bps, percent - into
   `meta.json`. Unit errors are the most expensive mistake available here, and
   they are silent.
5. **Every row records its source image.** No exceptions - `verification.md`
   depends on it and so does the sales trader.
6. **Do not tidy.** If the source says "PART", write "PART", not "POV".

## Verify

```bash
python .claude/skills/tca-review/scripts/verify.py --review reviews/<name>
```

The script checks what arithmetic can catch:

- venue or split shares summing to 100
- order counts summing to the stated total
- value summing to the stated total
- percentage-of-total columns summing to 100
- values outside a plausible range for their unit
- duplicate rows, missing strategies, orphan references
- sign convention against the stated convention in `meta.json`

Arithmetic cannot catch a transposition inside a single cell, so do the human
pass as well: **read each image again fresh, before looking at the CSV**, then
compare. Reading the CSV first makes you see what you wrote rather than what is
there.

Then write `verification.md`:

- every check, pass or fail
- every `NA`, named with its cell and its image
- every number you are less than certain of, named individually
- the units and the sign convention, stated plainly for confirmation

Show it to the user, ask for sign-off, and wait. On sign-off set `"verified":
true` in `meta.json`. `build_deck.py` refuses to run without it.

## Reading charts

Charts are for shape, not for digits. Take from a chart: the direction, the
ranking, the rough magnitude, whether the spread is wide or tight, where the
outliers sit.

Do not read precise values off a chart and present them as data. If a number
exists only in a chart, mark it approximate in the CSV (`approx` column = `Y`)
and say "about" wherever it appears.

Charts are also a cross-check: if the transcribed table disagrees with the shape
of the chart of the same thing, one of them is wrong. Find out which before
continuing.
