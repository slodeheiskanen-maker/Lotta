# Personnummer – last digit (check digit) in Excel

Open `personnummer_check_digit.xlsx`, paste numbers in column **A**, and the sheet works out the correct last digit, checks it, and shows gender.

## Swedish personnummer – one-cell formula

The number is in `A2`, written without its last digit as 9 digits (`YYMMDDNNN`, e.g. `811218987`):

```
=MOD(10-MOD(SUMPRODUCT(--MID(TEXT(MID(A2,{1,2,3,4,5,6,7,8,9},1)*{2,1,2,1,2,1,2,1,2},"00"),{1;2},1)),10),10)
```

Result: `6` → full number `811218-9876`.

For a full number like `19811218-9876` or `811218-9876`, put the 9 digits in a helper cell first:

```
B2: =SUBSTITUTE(SUBSTITUTE(TRIM(A2),"-",""),"+","")
C2: =CHOOSE(MATCH(LEN(B2),{9,10,11,12},0),B2,LEFT(B2,9),RIGHT(B2,9),MID(B2,3,9))
```

and use `C2` in the formula above. Valid if it equals `--RIGHT(B2,1)`.

## Finnish henkilötunnus – last character

`A2` = `131052-308T` (or without the last character):

```
=MID("0123456789ABCDEFHJKLMNPRSTUVWXY",MOD(--(LEFT(A2,6)&MID(A2,8,3)),31)+1,1)
```

## Notes
- Format column A as **Text**, otherwise Excel drops leading zeros (e.g. people born in 2000–2009).
- If your Excel uses `;` as the separator (Swedish/Finnish settings), replace `,` with `;` in the formulas.
  Inside `{...}`, also change the column separator to `\` (`{1\2\3...}`), and the row separator in `{1;2}` stays `;`.
