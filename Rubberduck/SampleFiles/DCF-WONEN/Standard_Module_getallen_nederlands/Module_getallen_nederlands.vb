Attribute VB_Name = "Module_getallen_nederlands"
Option Explicit

Option Compare Text

Dim eh$(99)
Dim vv$(12)

Function getaltekst(getal As Variant) As String
  Dim heel, deel    'decimal variants
  Dim txt$, n%      'string/int
  vulArrays
  heel = Int(CDec(Abs(getal)))
  deel = CDec(Abs(getal)) - heel

  txt = IIf(Sgn(getal) < 0, "min ", "") & _
    IIf(heel = 0, IIf(deel = 0, "nul", ""), spel(heel))

  If deel <> 0 Then
    txt = txt & IIf(heel = 0, "", " en ")
    n = Len(Mid(deel, 3))
    'boven miljoenste per macht van 3
    n = n + IIf(n < 6, 0, (3 - n Mod 3) Mod 3)
    deel = deel * (10 ^ n)
    txt = txt & spel(deel) & " " & _
      Trim(Replace(spel(10 ^ n), "??n", "")) & _
      IIf(n = 1, "de", "ste")
  End If

  getaltekst = txt
End Function



Function spel$(n)
  Dim t$, tmp$, B$, b1$, b2$
  Dim i%, s%, hv%, dv%

  t = CStr(n)
  'blokje van 4 bij getal tm 9999
  s = IIf(Len(t) = 4, 4, 3)
  'met nullen vullen tot lengte een veelvoud is van 3
  t = String((s - Len(t) Mod s) Mod s, "0") & t

  For i = 1 To Len(t) Step s
    tmp = Mid(t, i, s)
    b1 = Left(tmp, Len(tmp) - 2)
    hv = IIf(Right(b1, 1) = 0, 3, 2)    'duizend/honderd
    b1 = IIf(Right(b1, 1) = 0, Left(b1, 1), b1) 'idem

    b1 = xx(b1)
    b1 = IIf(b1 = "??n", " ", b1)       'geen ??nhonderd
    b1 = b1 & IIf(b1 = "", "", vv(hv))  'plak veelvoud

    b2 = Right(tmp, 2)
    dv = Len(t) - i - (s - 1)           'duizendvoud
    b2 = xx(b2)

    'spati?ring
    'optioneel EN voor getal tm 12
    b2 = IIf(dv = 0 And b1 <> "" And _
      Right(tmp, 2) > 0 And Right(tmp, 2) <= 12, _
      "en " & b2, b2)
    B = Trim(b1 & " " & b2) & " "
    'geen spatie veelvoud duizend hfdtelwoord tm honderd
    If (dv = 3 And Right(tmp, 2) = "00") Then B = Trim(B)
    'geen spatie veelvoud honderd
    If (dv = 3 And tmp < 100) Then B = Trim(B)

    spel = Trim(spel & " " & B & IIf(tmp = "000", "", vv(dv)))
  Next
End Function

Private Function xx$(n$)
'spelt tm 99
  If eh(n) <> "" Then
    xx = eh(n)
  Else
    xx = eh(Right(n, 1)) & _
      IIf(Left(n, 1) = 1 Or Right(n, 1) = 0, "", _
        IIf(Right(xx, 1) = "e", "?n", "en")) & _
      IIf(eh(Left(n, 1) * 10) <> "", eh(Left(n, 1) * 10), _
        eh(Left(n, 1)) & vv(1))
  End If
  xx = Trim(xx)
End Function

Private Sub vulArrays()
  eh(0) = " "
  eh(1) = "??n"
  eh(2) = "twee"
  eh(3) = "drie"
  eh(4) = "vier"
  eh(5) = "vijf"
  eh(6) = "zes"
  eh(7) = "zeven"
  eh(8) = "acht"
  eh(9) = "negen"
  eh(10) = "tien"
  eh(11) = "elf"
  eh(12) = "twaalf"
  eh(13) = "dertien"
  eh(14) = "veertien"
  eh(20) = "twintig"
  eh(30) = "dertig"
  eh(40) = "veertig"
  eh(80) = "tachtig"
  vv(1) = "tig"
  vv(2) = "honderd"
  vv(3) = "duizend"
  vv(6) = "miljoen"
  vv(9) = "miljard"
  vv(12) = "biljoen"
End Sub

Function GetalEuro(getal As Variant) As String
  Dim heel, deel    'decimal variants
  Dim txt$, n%      'string/int
  vulArrays
  heel = Int(CDec(Abs(getal)))
  deel = Round(CDec(Abs(getal)) - heel, 2)
  txt = IIf(Sgn(getal) < 0, "min ", "") & _
    IIf(heel = 0, IIf(deel = 0, "nul", ""), spel(heel))
  If heel > 0 Then txt = txt & IIf(deel = 0, " euro", " euro en ")
  If deel <> 0 Then
    deel = Int(Abs(deel * 100))
    txt = txt & spel(deel) & " cent"
  End If
  GetalEuro = txt
End Function

