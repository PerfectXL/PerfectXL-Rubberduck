Attribute VB_Name = "Module_gatallen_uitschrijven"
'Main Function
Function SpellNumber(ByVal MyNumber)
    Dim Dollars, Cents, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = " thousand "
    Place(3) = " million "
    Place(4) = " billion "
    Place(5) = " trillion "
    ' String representation of amount.
    MyNumber = Trim(Str(MyNumber))
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(MyNumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & _
                  "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(Right(MyNumber, 3))
        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Dollars
        Case ""
            Dollars = "no euros"
        Case "One"
            Dollars = "one euro"
         Case Else
            Dollars = Dollars & " euros"
    End Select
    Select Case Cents
        Case ""
            Cents = " "
        Case "One"
            Cents = " and one cent"
              Case Else
            Cents = " and " & Cents & " cents"
    End Select
    SpellNumber = Dollars & Cents
End Function
      
' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " hundred "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function
      
' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "ten"
            Case 11: Result = "eleven"
            Case 12: Result = "twelve"
            Case 13: Result = "thirteen"
            Case 14: Result = "fourteen"
            Case 15: Result = "fifteen"
            Case 16: Result = "sixteen"
            Case 17: Result = "seventeen"
            Case 18: Result = "eighteen"
            Case 19: Result = "nineteen"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "twenty "
            Case 3: Result = "thirty "
            Case 4: Result = "forty "
            Case 5: Result = "fifty "
            Case 6: Result = "sixty "
            Case 7: Result = "seventy "
            Case 8: Result = "eighty "
            Case 9: Result = "ninety "
            Case Else
        End Select
        Result = Result & GetDigit _
            (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = Result
End Function
     
' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "one"
        Case 2: GetDigit = "two"
        Case 3: GetDigit = "three"
        Case 4: GetDigit = "four"
        Case 5: GetDigit = "five"
        Case 6: GetDigit = "six"
        Case 7: GetDigit = "seven"
        Case 8: GetDigit = "eight"
        Case 9: GetDigit = "nine"
        Case Else: GetDigit = ""
    End Select
End Function








