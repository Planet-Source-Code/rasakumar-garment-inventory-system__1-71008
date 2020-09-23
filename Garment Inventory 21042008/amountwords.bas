Attribute VB_Name = "Others"
    Dim str As String
    Dim part(3)
    Dim part2(3)
    Dim rsuserlog As New ADODB.Recordset
    Public Word As String
    Public Text(28) As String
Public Function towords(Value As Double) As String
  Dim Ones As Single, Tens As Single, Huns As Single, Ths As Single, Lakhs As Single
  Word = ""
  Text(0) = ""
  Text(1) = "One"
  Text(2) = "Two"
  Text(3) = "Three"
  Text(4) = "Four"
  Text(5) = "Five"
  Text(6) = "Six"
  Text(7) = "Seven"
  Text(8) = "Eight"
  Text(9) = "Nine"
  Text(10) = "Ten"
  Text(11) = "Eleven"
  Text(12) = "Twelve"
  Text(13) = "Thirteen"
  Text(14) = "Fourteen"
  Text(15) = "Fifteen"
  Text(16) = "Sixteen"
  Text(17) = "Seventeen"
  Text(18) = "Eighteen"
  Text(19) = "Ninteen"
  Text(20) = "Twenty"
  Text(21) = "Thirty"
  Text(22) = "Forty"
  Text(23) = "Fifty"
  Text(24) = "Sixty"
  Text(25) = "Seventy"
  Text(26) = "Eighty"
  Text(27) = "Ninty"
  Text(28) = "Hundred"
  If Value <= 20 Then
    PrintOne (Value)
  ElseIf Value > 20 And Value < 100 Then
    PrintTens (Value)
  ElseIf Value >= 100 And Value < 1000 Then
    PrintHuns (Value)
  ElseIf Value >= 1000 And Value < 100000 Then
    PrintThs (Value)
  ElseIf Value >= 100000 And Value < 10000000 Then
    PrintLakhs (Value)
  ElseIf Value >= 10000000 Then
    PrintCrore (Value)
  End If
  towords = Word & " Only"
End Function
Private Sub PrintOne(Val As Integer)
  Word = Word & Text(Val)
End Sub
Private Sub PrintTens(Val As Integer)
  If Val > 20 Then
    Word = Word & Text(18 + Int(Val / 10)) & " " & Text(Val Mod 10)
  Else
    PrintOne (Val)
  End If
End Sub
Private Sub PrintHuns(Val As Integer)
  Word = Word & Text(Int(Val / 100)) & " Hundred "
  Val = Val Mod 100
  If Val <> 0 Then
    Word = Word & "and "
    PrintTens (Val)
  End If
End Sub

Private Sub PrintThs(Val As Single)

  PrintTens (Int(Val / 1000))
  Word = Word & " Thousand "
  Val = Val Mod 1000
  If Val >= 100 Then
    PrintHuns (Val)
  Else
    If Val <> 0 Then Word = Word & "and "
    PrintTens (Val)
  End If

End Sub

Private Sub PrintLakhs(Val As Single)

  PrintTens (Int(Val / 100000))
  Word = Word & " Lakh "
  Val = Val Mod 100000
  If Val >= 1000 Then
    PrintThs (Val)
  ElseIf Val >= 100 Then
    PrintHuns (Val)
  Else
    If Val <> 0 Then Word = Word & "and "
    PrintTens (Val)
  End If

End Sub

Private Sub PrintCrore(Val As Double)

  PrintTens (Int(Val / 10000000))
  Word = Word & " Crore "
  Val = Val Mod 10000000
  If Val >= 100000 Then
    PrintLakhs (Val)
  ElseIf Val >= 1000 Then
    PrintThs (Val)
  ElseIf Val >= 100 Then
    PrintHuns (Val)
  Else
    If Val <> 0 Then Word = Word & "and "
    PrintTens (Val)
  End If

End Sub
