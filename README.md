# A and P

Mars infrastructure catalog.

**A** is the algorithm.  
**P** is the pound index and the packet ceiling.

Unit of catalog mass is one pound.  
Index rule: **n ≠ 0**. **n = 1 + 0**.

This is a *could-go* catalog. It lists infrastructure that could go to Mars. It is not a flight manifest and not a promise that any row will fly.

## Clone

```bash
git clone https://github.com/fitzyracing1/A-and-P.git
cd A-and-P
```

Or unpack the release zip and work offline.

## What is in this repo

| File | What it is |
|---|---|
| `A-and-P-catalog-matrix.xlsx` | Working matrix (Algorithm, Catalog, Category matrix, How to add) |
| `catalog.csv` | Flat export of filled catalog rows n = 1 … 50 |
| `ALGORITHM.md` | The n = 1 + 0 rule |
| `build_matrix.py` | Regenerates the workbook from seed data |

## How A works

- First legal index is **n = 1**.
- Offset is **0**.
- Generator: `n_start = 1 + 0`.
- New items append at `n_max + 1`.
- Retired items keep their number. They are not reused as 0.

## How P works

- `m_i` = catalog pounds for item i (decimals allowed).
- `p_i` = packet ceiling = `CEILING(m_i, 1)` if you later packetize.
- Catalog itself does not require physical 1-lb crates.

## Could-go classes

- **Core** — needed for a minimal useful site
- **Enabling** — unlocks Core so it actually works
- **Optional** — science, comfort, extra margin

## Add a row

1. Next empty n on the Catalog sheet is already numbered (51–60 reserved).
2. Fill Category, Item, Catalog lb (blue), Assembly k, class, depends, enables, notes.
3. Do not put anything at index 0. Do not renumber existing n downward.
4. Depends-on uses other n values (example: `5,6`). Use `—` if it depends on nothing.

## Rebuild the workbook

```bash
python3 build_matrix.py
```

Requires `openpyxl`. Output path inside the script points at the artifacts copy; edit the output path if you run it from a clone.

## License

MIT. See `LICENSE`.
