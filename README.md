# Personnummer – last digit (check digit) in Excel

Open `personnummer_check_digit.xlsx`, paste numbers in column **A**, and the sheet works out the correct last digit, checks it, and shows gender.

## Suomenkielinen Excel (Finnish Excel)

Kaavat suomenkielisille Excel-funktioille. Erotin on `;`.

### Henkilötunnuksen tarkistusmerkki

`A2` = `131052-308` tai `131052-308T`:

```
=POIMI.TEKSTI("0123456789ABCDEFHJKLMNPRSTUVWXY";JAKOJ.ANNOS(--(VASEN(A2;6)&POIMI.TEKSTI(A2;8;3));31)+1;1)
```

Tulos: `T`. Tarkistus (onko viimeinen merkki oikein):

```
=JOS(OIKEA(A2;1)=POIMI.TEKSTI("0123456789ABCDEFHJKLMNPRSTUVWXY";JAKOJ.ANNOS(--(VASEN(A2;6)&POIMI.TEKSTI(A2;8;3));31)+1;1);"OK";"VÄÄRIN")
```

Sukupuoli: `=JOS(ONPARITON(--POIMI.TEKSTI(A2;10;1));"Mies";"Nainen")`

### Ruotsalaisen personnummerin viimeinen numero

`A2` = 9 numeroa ilman viivaa ja viimeistä numeroa (esim. `811218987`):

```
=JAKOJ.ANNOS(10-JAKOJ.ANNOS(TULOJEN.SUMMA(POIMI.TEKSTI(A2;RIVI(EPÄSUORA("1:9"));1)*(2-JAKOJ.ANNOS(RIVI(EPÄSUORA("1:9"))+1;2))-9*(POIMI.TEKSTI(A2;RIVI(EPÄSUORA("1:9"));1)*(2-JAKOJ.ANNOS(RIVI(EPÄSUORA("1:9"))+1;2))>9));10);10)
```

Tulos: `6` → `811218-9876`.

Muotoile sarake A **tekstiksi**, jotta etunollat säilyvät.

---

# English Excel

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
- Using a non-English Excel? See the Finnish section above. The workbook itself works in any language, because Excel translates formulas automatically.
