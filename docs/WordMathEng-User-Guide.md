# WordMathEng user guide

WordMathEng works in two kinds of equation, which behave the same way:

- **Word equations** (equation editor, Alt+=): click inside the equation and use the ribbon's calculate
  (**=**) or **Solve** button.
- **Free-typed equations** in ordinary text, written between `~~` markers: `~~F = m·a = ?~~`. Put the
  cursor inside the markers and press calculate, or run the macro `WM_EvaluateAllFreeText` (Alt+F8) to
  calculate the whole document from top to bottom.

Numbers always use a **decimal point** (`2.5`). Lists and arrays separate their elements with `;`.

## Free-typed equations

| You type | Meaning |
|---|---|
| `~~m := 5 [kg]~~` | a definition (`:=` or `≔`); applies to everything below it, in free text and in Word equations |
| `~~F = m·a = ?~~` | calculate: the answer replaces the `?` |
| `~~F := m·a = ? [kN]~~` | define and calculate in one: defines `F` and writes its value after it (`= 2.5 [kN]`); recalculated in place like any answer. Word equations: `F≔m·a=?` |
| `~~F = m·a = ? [kN]~~` | calculate in a chosen unit |
| `~~F = m·a~~` | a plain equation (a relation); not calculated, but used for solving |
| `~~a = ?~~` | solve: rearranges the nearest earlier equation that contains `a` and whose other symbols are defined |
| `~~F = m·a = ? #(3)~~` | an equation number, kept as it is |

- An equation must start and end in the same paragraph. Text outside the markers is ignored, and a single `~`
  is ordinary text.
- **Before printing or publishing:** click **Format for print** in the ribbon's *Free-typed ~~* group. Every
  `~~` marker becomes white, 1 pt (character style "WM Marker Print"), and the equations still calculate.
  **Show markers** returns them to normal. One Ctrl+Z also undoes it.
  - White text still comes along when copying, and shows on a coloured background or highlighting.
- **Hiding the markers instead:** run the macro `WM_HideMarkers` (Alt+F8). The markers get the character
  style "WM Marker" with hidden text, so they disappear when ¶ is off and don't print. **Show markers**
  undoes this too.
- **Recalculation:** pressing calculate on an equation that already has an answer recalculates it and
  replaces the answer, keeping its unit.
  - This includes solved answers such as `~~a = 2 [m/s^2]~~`.
  - Note: an equation you typed yourself in the form `name = number`, e.g. `~~x = 5~~`, is also recalculated
    if something gives `x` a value.
- The **Solve** button inside `~~ … ~~` explains the `~~a = ?~~` syntax.
- **Define and ask:** the definition part (`F := m·a`) applies to later equations whether or not the answer has been written yet; the answer after it is ignored. A function definition (`~~f(x) := x^2 = ?~~`) is defined but gets no value, with a message.

## Units

Units are written in **square brackets**, after a number or an expression: `5 [kg]`, `9.81 [m/s^2]`,
`2 [kN·m]`. Names outside brackets are always variables, so `m` (a mass) and `[m]` (metres) never clash.

- **Case matters:** `[mPa]` is not `[MPa]`, and `[Bar]` is an error ("did you mean bar?").
- Inside brackets, `·`, `*` or `.` multiply (`[lb.ft]`, `[kN.m]`), `/` divides and `^` raises to a power.
- `[ ]` are reserved for units and arrays, so use `( )` for grouping.
- **Dimension check:** `5 [m] + 3 [s]` gives a unit mismatch error. So does asking for a result in a unit of
  the wrong dimension.

### Which unit the answer is given in (`= ?` without a unit)

1. Units that cancel are cancelled: `[bar]·[m^2]` → `[N]`, `[N]/[kg]` → `[m/s^2]`.
2. Combinations of base units become a named unit: `[kg]·[m/s^2]` → `[N]`.
3. Combinations of named units that don't cancel are **never renamed**: `[kN]·[m]` stays `[kN·m]` (not
   `[kJ]`), and `[N]/[mm^2]` stays `[N/mm^2]` (not `[MPa]`). Ask for `= ? [MPa]` if you want MPa.
4. A prefix is chosen automatically: `1 [bar]·1 [m^2]` → `100 [kN]`.
5. If all the inputs are imperial, so is the answer (psi, lbf, ft, lb…). If they are mixed, the answer is SI.
6. A sum is given in its largest input unit: `2 [kN] + 500 [N]` → `2.5 [kN]`.

### Imperial and US customary units

`[in]`, `[ft]`, `[yd]`, `[mi]`, `[lb]` (pound-mass), `[lbf]` (pound-force), `[kip]`, `[psi]`, `[ksi]`,
`[slug]`, `[oz]`, `[mph]`, `[hp]`, `[BTU]`, and `[lb.ft]` / `[lbf.ft]`.

- `[lb]` is a mass. If it is used where a force is needed, the error suggests `[lbf]`.
- Units that differ between the US and the UK must be explicit: `[gal_US]` / `[gal_UK]` and
  `[ton_short]` / `[ton_long]` (`[t]` is the metric tonne).

### Temperatures and gauge pressures

| Unit | Meaning |
|---|---|
| `[°C]`, `[°F]` (or `[degC]`, `[degF]`) | absolute temperatures |
| `[Δ°C]`, `[Δ°F]` (or `[DdegC]`, `[DdegF]`) | temperature differences |
| `[K]`, `[°R]` (`[degR]`) | absolute (no offset) |
| `[barg]`, `[psig]` | gauge pressure (absolute − 101.325 kPa) |
| `[bara]`, `[psia]`, `[bar]`, `[psi]`, `[Pa]` | absolute pressure, or a pressure difference or stress |

- absolute − absolute gives a difference: `30 [°C] − 20 [°C]` → `10 [Δ°C]`;
- absolute + difference stays absolute: `20 [°C] + 5 [Δ°C]` → `25 [°C]`;
- absolute + absolute gives a warning;
- a gauge pressure multiplied or divided by another quantity uses the gauge value:
  `2 [barg]·1 [m^2]` → `200 [kN]`.

Absolute units can't be used in an equation that is *solved* for an unknown; use `[K]` or `[bara]` there.

## Arrays (free text and Word equations)

```
~~a := [1; 2; 3] [kN]~~
~~F = 2·a = ?~~                 →  F = 2·a = [2; 4; 6] [kN]
~~a = ? [N]~~                   →  a = [1000; 2000; 3000] [N]
~~b := [1; 4; 9]~~
~~sqrt(b) = ?~~                 →  sqrt(b) = [1; 2; 3]
```

- Everything is calculated **element by element**. A single number combines with every element.
- Two arrays in one expression must have the same length.
- An array has **one unit** for all its elements. The prefix is chosen from the largest element.
- Arrays can be **solved**: with `~~sa := [1; 2; 3]~~`, `~~sb := [4; 5; 6]~~` and `~~sa = sb + sc~~`,
  `~~sc = ?~~` gives `[-3; -3; -3]`.
- Arrays of absolute units (`[°C]`, `[barg]`) and arrays of arrays are not supported. Use a Word equation
  matrix for 2-D data.

## Matrices (Word equations only)

Type matrices with the equation editor's matrix structures. WordMat's matrix maths works:

- `+`, `−`, scalar ×, matrix product, inverse `A^(-1)`, `det(A)`, `transpose(A)`;
- solving `A·x = b` with the Solve button;
- vector dot and cross products.

Matrices can be defined (`A≔(■(…))`) and may carry **one unit** for all elements:
`σ≔(■(80&30@30&−40)) [MPa]`.

### Principal stresses: `eigenval`, `eigenvec`, `principalangle`

```
σ≔(■(80&30@30&−40)) [MPa]
eigenval(σ)=?          →  (87.08; −47.08) [MPa]      sorted σ1 ≥ σ2 (≥ σ3)
eigenvec(σ)=?          →  unit eigenvectors as columns, in the same order
principalangle(σ)=?    →  13.28°                      2×2 only: angle of σ1 from the x axis
eigenval(σ)=? [kPa]    →  in a chosen unit
```

- These work for 2×2 and 3×3 matrices, using the Jacobi method for symmetric matrices.
- A non-symmetric matrix gives a warning and uses Maxima's general routine.
- Exact or decimal output follows WordMat's exact/numeric setting. Results with a unit are always decimals.
- `principalangle` follows WordMat's degrees/radians setting.

## Settings worth knowing

- **Decimal point only.** WordMat's decimal-comma option has been removed.
- WordMat's own "units mode" (unbracketed units, on the ribbon) still exists for compatibility. Square-bracket
  units work whether it is on or off.
- WordMathEng never contacts the internet to check for updates.
