# funny-bingo

Config-driven Bingo ticket generator for Microsoft Word (.docx), designed for print-ready layout control with `python-docx`.

## Features

- Supports **3x3, 5x5, and 7x7** grids with a `⭐` FREE cell in the center
- Generates `rounds * tickets_per_round` tickets (e.g., `3 * 170 = 510`)
- Ticket structure:
	- Header row 1: `Round <n>`
	- Header row 2: configurable title
	- N×N game grid with `⭐` in center
	- Footer row with merged special task + ticket ID
- Enforces no duplicate criteria within a ticket
- Enforces unique ticket grids across all generated tickets
- Page layout via parent container table:
	- `tickets_per_row` x `rows_per_page`
	- e.g., `3 x 3 = 9 tickets/page` for 5x5
- Fixed cell width/height for deterministic printing
- Configurable page orientation and margins (landscape + small margins recommended)
- Optional per-round output files
- Shared data pool via `--data` flag (`tasks`, `criteria_pool`)

## Files

- `generate_bingo_docx.py` — main generator script
- `data.json` — shared pool of tasks and criteria used by all configs
- `config.sample.json` — sample config with inline tasks and criteria (self-contained)
- `config_3x3.json` — ready-to-use 3×3 config (1 round, 170 tickets)
- `config_5x5.json` — ready-to-use 5×5 config (3 rounds, 100 tickets)
- `config_7x7.json` — ready-to-use 7×7 config (3 rounds, 170 tickets)
- `requirements.txt` — dependencies

## Install

```bash
pip install -r requirements.txt
```

## Run

### Shortcut (auto-loads `config_{N}x{N}.json` + `data.json`)

```bash
python generate_bingo_docx.py --grid 5
python generate_bingo_docx.py --grid 7 --output bingo_7x7.docx
```

### Full form

```bash
python generate_bingo_docx.py --config config_5x5.json --data data.json --output bingo_tickets.docx
```

### Optional flags

```bash
# Reproducible output
python generate_bingo_docx.py --grid 5 --seed 42

# Also export one file per round
python generate_bingo_docx.py --grid 7 --per-round-files
```

## Config Notes

- `grid_size` must be odd (for a center FREE cell); supported values: `3`, `5`, `7`
- `criteria_pool` must have at least `grid_size * grid_size - 1` unique items
- `tickets_per_page` must be divisible by `tickets_per_row`
- `page_size` supports `A4` and `LETTER` (default `A4`)
- `orientation` supports `portrait` or `landscape` (default `landscape`)
- `margin_cm` controls all page margins (small print margins recommended)
- `header_font_size` default `16`
- `header_font_style` default `all-caps`
- `header_bg_color` default `grey` (or hex like `#808080`)
- `header_color` default `white` (or hex like `#FFFFFF`)
- `round_height_cm` default `0.5`
- `header_height_cm` default `0.5`
- `footer_height_cm` default `0.3`
- `text_font` default `Calibri`
- `text_size` default `12`
- `footer_font_size` default `10`
- The script validates whether configured ticket dimensions fit the current Word page printable area and fails fast if they do not
- `tasks` and `criteria_pool` can be embedded in the config or loaded separately via `--data data.json`

## Output

- Main file: `bingo_tickets.docx` (or named via `--output`)
- With `--per-round-files`:
	- `bingo_tickets_round_1.docx`
	- `bingo_tickets_round_2.docx`
	- etc.
