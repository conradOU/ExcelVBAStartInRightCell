Private Sub Workbook_Open()

    Dim todayIs As Integer
    todayIs = Weekday(Date, vbMonday)

    Select Case todayIs
       Case 1
          Range("B3").Select
       Case 2
          Range("C3").Select
       Case 3
          Range("D3").Select
       Case 4
          Range("E3").Select
       Case 5
          Range("F3").Select
       Case 6
          Range("G3").Select
       Case 7
          Range("H3").Select
       End Select

End Sub
