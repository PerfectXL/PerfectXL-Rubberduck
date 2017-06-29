Attribute VB_Name = "Module1"
Sub addPenalty()
    If [isGameOn] Then
        [pCount] = [pCount] + 1
    End If
End Sub
Sub startGame()
    Application.ScreenUpdating = False
    Sheets("Second").Visible = True
    Sheets("Second").Select
    Range("hetch14").Select
    ResetGame
    Sheets("First").Select
    Range("startCell").Select
    Application.ScreenUpdating = True
    [endTime] = ""
    [pCount] = 0
    [isGameOn] = True
    [startTime] = Now
End Sub
Sub checkCell(ByVal target As Range)
    Debug.Print target.address
    
    If [isGameOn] = True Then
        'check if any of the 21 of cells are visited and mark them done
        On Error Resume Next
        Dim nameOfRange As String
        
        nameOfRange = ""
        nameOfRange = target.Name.Name
        If nameOfRange <> "" Then
            Range("lstVisited").Cells(Application.WorksheetFunction.Match(nameOfRange, Range("lstNames"), 0)).Value = True
        End If
    End If
End Sub
Sub resetVisitedCells()
    Range("lstVisited").Value = ""
End Sub
Sub ResetGame()
    [endTime] = ""
    [pCount] = 0
    [isGameOn] = False
    [startTime] = ""
    resetVisitedCells
    Sheets("Results").Visible = False
    Sheets("Second").Visible = False
    Range("hetch10").Value = "Select x from drop down"
    DoEvents
End Sub
Sub showTime()
    Range("timeElapsed").Calculate
    DoEvents
End Sub
Sub twitterLink()
    Range("B13").Select
End Sub
Sub fbLink()
    Range("D13").Select
End Sub

