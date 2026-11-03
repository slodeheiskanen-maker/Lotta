Attribute VB_Name = "Henkilotunnus"
Option Explicit

' =====================================================================
'  Helppoja funktioita henkilötunnukselle ja ruotsalaiselle personnummerille
'  Easy functions for Finnish henkilötunnus and Swedish personnummer
'
'  Suomi / Finland:
'    =HETU_TARKISTE(A2)    -> oikea viimeinen merkki, esim. "T"
'    =HETU_OIKEIN(A2)      -> TOSI / EPÄTOSI (TRUE / FALSE)
'    =HETU_TAYSI(A2)       -> koko tunnus oikealla tarkisteella, esim. "131052-308T"
'    =HETU_SUKUPUOLI(A2)   -> "Mies" / "Nainen"
'
'  Ruotsi / Sweden:
'    =PNR_TARKISTE(A2)     -> oikea viimeinen numero, esim. 6
'    =PNR_OIKEIN(A2)       -> TOSI / EPÄTOSI
'    =PNR_TAYSI(A2)        -> esim. "811218-9876"
'
'  Tunnus voi olla viimeisen merkin kanssa tai ilman.
'  The number can be given with or without its last character.
' =====================================================================

Private Const HETU_MERKIT As String = "0123456789ABCDEFHJKLMNPRSTUVWXY"

' ---------- Suomi: henkilötunnus ----------

Private Function HetuSiivo(ByVal tunnus As Variant) As String
    HetuSiivo = UCase$(Trim$(CStr(tunnus)))
End Function

Public Function HETU_TARKISTE(ByVal tunnus As Variant) As Variant
    Dim s As String, numero As String
    s = HetuSiivo(tunnus)
    If Len(s) < 10 Then HETU_TARKISTE = CVErr(xlErrValue): Exit Function
    numero = Left$(s, 6) & Mid$(s, 8, 3)
    If Not IsNumeric(numero) Or InStr(numero, ",") > 0 Or InStr(numero, ".") > 0 Then
        HETU_TARKISTE = CVErr(xlErrValue): Exit Function
    End If
    HETU_TARKISTE = Mid$(HETU_MERKIT, (CLng(numero) Mod 31) + 1, 1)
End Function

Public Function HETU_OIKEIN(ByVal tunnus As Variant) As Boolean
    Dim s As String, t As Variant
    s = HetuSiivo(tunnus)
    If Len(s) <> 11 Then Exit Function
    t = HETU_TARKISTE(s)
    If IsError(t) Then Exit Function
    HETU_OIKEIN = (Right$(s, 1) = t)
End Function

Public Function HETU_TAYSI(ByVal tunnus As Variant) As Variant
    Dim t As Variant
    t = HETU_TARKISTE(tunnus)
    If IsError(t) Then HETU_TAYSI = t: Exit Function
    HETU_TAYSI = Left$(HetuSiivo(tunnus), 10) & t
End Function

Public Function HETU_SUKUPUOLI(ByVal tunnus As Variant) As Variant
    Dim s As String
    s = HetuSiivo(tunnus)
    If Len(s) < 10 Or Not IsNumeric(Mid$(s, 10, 1)) Then
        HETU_SUKUPUOLI = CVErr(xlErrValue): Exit Function
    End If
    If CInt(Mid$(s, 10, 1)) Mod 2 = 1 Then
        HETU_SUKUPUOLI = "Mies"
    Else
        HETU_SUKUPUOLI = "Nainen"
    End If
End Function

' ---------- Ruotsi: personnummer ----------

' Palauttaa vain numerot / digits only
Private Function VainNumerot(ByVal tunnus As Variant) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(CStr(tunnus))
        c = Mid$(CStr(tunnus), i, 1)
        If c Like "#" Then r = r & c
    Next i
    VainNumerot = r
End Function

' Ottaa 9 numeroa VVKKPPNNN / takes the 9 digits YYMMDDNNN
Private Function PnrYhdeksan(ByVal tunnus As Variant) As String
    Dim s As String
    s = VainNumerot(tunnus)
    Select Case Len(s)
        Case 9:  PnrYhdeksan = s
        Case 10: PnrYhdeksan = Left$(s, 9)
        Case 11: PnrYhdeksan = Right$(s, 9)
        Case 12: PnrYhdeksan = Mid$(s, 3, 9)
        Case Else: PnrYhdeksan = ""
    End Select
End Function

Public Function PNR_TARKISTE(ByVal tunnus As Variant) As Variant
    Dim s As String, i As Long, p As Long, summa As Long
    s = PnrYhdeksan(tunnus)
    If s = "" Then PNR_TARKISTE = CVErr(xlErrValue): Exit Function
    For i = 1 To 9
        p = CLng(Mid$(s, i, 1)) * IIf(i Mod 2 = 1, 2, 1)
        If p > 9 Then p = p - 9
        summa = summa + p
    Next i
    PNR_TARKISTE = (10 - (summa Mod 10)) Mod 10
End Function

Public Function PNR_OIKEIN(ByVal tunnus As Variant) As Boolean
    Dim s As String, t As Variant
    s = VainNumerot(tunnus)
    If Len(s) <> 10 And Len(s) <> 12 Then Exit Function
    t = PNR_TARKISTE(s)
    If IsError(t) Then Exit Function
    PNR_OIKEIN = (CLng(Right$(s, 1)) = t)
End Function

Public Function PNR_TAYSI(ByVal tunnus As Variant) As Variant
    Dim s As String, t As Variant
    s = PnrYhdeksan(tunnus)
    t = PNR_TARKISTE(tunnus)
    If IsError(t) Then PNR_TAYSI = t: Exit Function
    PNR_TAYSI = Left$(s, 6) & "-" & Right$(s, 3) & t
End Function
