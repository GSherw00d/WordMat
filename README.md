# WordMathEng

**WordMathEng** is a maths add-in for Microsoft Word on Windows, aimed at engineering calculations. It is a
modified version of [WordMat](https://github.com/Eduap-com/WordMat) by Eduap and uses the
[Maxima](https://maxima.sourceforge.io/) computer algebra system for all calculations.

Results are written into the document as ordinary, editable Word text or equations, never as images.

> **Status:** in active development. It is built and tested on Windows desktop Word (Microsoft 365) only.
> There is no installer yet.

## What it adds to WordMat

| Feature | Example |
|---|---|
| **Units** in square brackets, with dimension checking | `~~1 [bar]·1 [m^2] = ?~~` → `100 [kN]`; `5 [m] + 3 [s]` → unit mismatch error |
| **Output units** per equation | `~~1 [psi] = ? [kPa]~~` → `6.894757 [kPa]` |
| **Imperial / US units** | `[psi]`, `[ksi]`, `[lbf]`, `[lb]`, `[lb.ft]`, `[ft]`, `[in]`, `[gal_US]`, `[gal_UK]` … |
| **Absolute and difference temperatures, gauge pressures** | `20 [°C] + 5 [Δ°C]` → `25 [°C]`; `2 [barg]·1 [m^2]` → `200 [kN]` |
| **Free-typed equations** in ordinary text | `~~F = m·a = ?~~` |
| **Solving for an unknown** in free text | `~~F = m·a~~` then `~~a = ?~~` |
| **Recalculation in place** | pressing calculate again updates an existing answer |
| **Arrays** with one unit | `~~a := [1; 2; 3] [kN]~~`, `~~2·a = ?~~` → `[2; 4; 6] [kN]` |
| **Principal stresses** | `eigenval(σ)`, `eigenvec(σ)`, `principalangle(σ)` |
| **Safe equation reading** | the document and its undo history are never changed to read an equation |

See the **[user guide](docs/WordMathEng-User-Guide.md)** for the full syntax.

## Changes from upstream WordMat

- Renamed to WordMathEng. Internal names (registry key `HKCU\Software\WordMat`, the `%AppData%\WordMat`
  folder, VBA module names) are unchanged.
- **No links to the paid WordMat products:** the WordMat+ menu, partnership information, online update checks
  and downloads, and online help hosted by Eduap have been removed. WordMathEng never contacts wordmat.dk or
  eduap.com.
- **Decimal point only:** the decimal-comma setting has been removed; `;` separates list and array elements.
- Start-up no longer changes Word's global settings (maths AutoCorrect, AutoFormat-as-you-type).
- Bug fixes:
  - `det(■(…))`, `transpose(■(…))` and `eigenvalues(■(…))` now work with the matrix inside the function's brackets;
  - `1/0` is reported as an error instead of writing `≈fejl`;
  - reading a bold equation no longer un-bolds it.
- The ribbon is kept as text in `Windows/customUI/customUI14.xml`.

## Source layout

| Path | Contents |
|---|---|
| `Windows/VBA-modules/` | VBA source (`.bas`, `.cls`, `.frm`), Windows-1252 with CRLF |
| `Windows/customUI/customUI14.xml` | the ribbon |
| `Shared/Maxima-files/` | Maxima code, including the unit engine (`WMUnits.mac`, `WMUnitTable.mac`) and matrix functions (`WMMatrix.mac`) |
| `Shared/Translations/win/` | user-interface texts |
| `Windows/WordMat.dotm` | upstream's template, used as the seed for builds |

New modules written for WordMathEng:
- `WMFreeText.bas`: free-typed equations;
- `WMUnits.bas`: unit brackets;
- `WMSafeRead.bas`: non-destructive equation reading;
- `WMTest.bas`: automated test entry points.

The build scripts (VBA source → `WordMathEng.dotm`) and the automated test suite currently live in a
separate project repository that is not yet published. Without them, the modules can be imported into a copy
of `Windows/WordMat.dotm` from Word's VBA editor.

## Licence

WordMathEng, like WordMat's core, is free software under the **GNU General Public License v3 or later**. See
[LICENSE](LICENSE). It is based on WordMat by Eduap (Mikael Samsøe Sørensen) and does not include the
proprietary WordMat components (installers, WordMat+/WordMatP).

Maxima is © its authors and licensed under the GNU GPL.
