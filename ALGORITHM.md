# A and P — algorithm

## Name

- **A** — algorithm that addresses every catalog row
- **P** — pound as the catalog unit, and `p_i` as packet ceiling

## Generator

```
n ≠ 0
n = 1 + 0
```

| Symbol | Meaning | Value |
|---|---|---|
| n | Catalog index | 1, 2, 3, … |
| n0 | Offset | 0 |
| n_start | First legal index | 1 + 0 = 1 |
| m_i | Catalog pounds of item i | input |
| p_i | Packet ceiling | CEILING(m_i, 1) |
| k_i | Assembly width | input |
| class_i | Could-go class | Core / Enabling / Optional |

## Index rules

1. There is no row 0 and no item 0. Empty is not an index.
2. Row i on the Catalog sheet is addressed as n = i, with i starting at 1.
3. New items append at n_max + 1.
4. Deleting an item retires the number. It is not reused as 0.
5. This workbook catalogs infrastructure that *could* go to Mars. It is not a flight list.

## Why not zero

Zero would mean “no item” and “first item” in the same slot. A and P keep those apart. The first real piece of infrastructure is n = 1. The offset stays 0 so the generator is explicit: start at one, add nothing, never decrement to zero.
