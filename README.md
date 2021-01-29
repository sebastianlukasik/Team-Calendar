# Team-Calendar
Projekt przy użyciu języka Visual Basic. Projekt miał na celu stworzenie pliku, który może służyć jako grafik na kilka lat. User Form, który pojawia się na starcie pliku pozwala na dodanie nowego pracownika oraz na administrowanie i zarządzanie grafikiem. 


Sub oneMonth(iColumn As Integer, dtData As Date, iMonth As Integer)
Dim iWiersz, iWeekDay, iWierszTeam As Integer
Dim strWeekday, sMember As String

iWiersz = 2

Do While Month(dtData) = iMonth
    ThisWorkbook.Sheets("Kalendarz").Cells(iWiersz, iColumn) = dtData
    iWeekDay = Weekday(dtData, vbMonday)
    strWeekday = WeekdayName(iWeekDay, False, vbMonday)
    ThisWorkbook.Sheets("Kalendarz").Cells(iWiersz, iColumn + 1) = strWeekday
    
    If iWeekDay < 6 Then
    iWierszTeam = 2
    Do Until ThisWorkbook.Sheets("Team").Cells(iWierszTeam, "B") = iWeekDay
        iWierszTeam = iWierszTeam + 1
    Loop
    sMember = ThisWorkbook.Sheets("Team").Cells(iWierszTeam, "A")
    Else
    sMember = "Wolne"
    End If
    ThisWorkbook.Sheets("Kalendarz").Cells(iWiersz, iColumn + 2) = sMember
    
    
    dtData = DateAdd("d", 1, dtData)
    iWiersz = iWiersz + 1
Loop
     
dtData = DateAdd("m", -1, dtData)


End Sub

Sub Calendar()
Dim iCalendarMonth As Integer
Dim iCalendarColumn As Integer
Dim dtCalendarDate As Date
Dim iYear As Integer
Dim iQuestion As Integer


iYear = InputBox("Wprowadź rok:", "Grafik")
iQuestion = MsgBox("Czy chcesz wyświetlić grafik dla " & iYear & "?", vbYesNo + vbQuestion, "Grafik")
If iQuestion = vbYes Then
    GoTo continue
Else
    End
End If
continue:

If IsNumeric(iYear) = False Then End
iCalendarMonth = 1
iCalendarColumn = 1
dtCalendarDate = Format("01.01." & iYear, "dd.mm.yyyy")
Do While iCalendarMonth < 7 ' Untill iCalendarMonth=7
    Call oneMonth(iCalendarColumn, dtCalendarDate, iCalendarMonth)
    iCalendarColumn = iCalendarColumn + 3
    iCalendarMonth = iCalendarMonth + 1
    dtCalendarDate = DateAdd("m", 1, dtCalendarDate)
Loop
Call MsgBox("Odświeżanie ukończone:", vbOKOnly + vbInformation, "Grafik")
End Sub


#Tworzenie User Forms

Sub oneMonth(iColumn As Integer, dtData As Date, iMonth As Integer)
Dim iWiersz, iWeekDay, iWierszTeam As Integer
Dim strWeekday, sMember As String

iWiersz = 2

Do While Month(dtData) = iMonth
    ThisWorkbook.Sheets("Kalendarz").Cells(iWiersz, iColumn) = dtData
    iWeekDay = Weekday(dtData, vbMonday)
    strWeekday = WeekdayName(iWeekDay, False, vbMonday)
    ThisWorkbook.Sheets("Kalendarz").Cells(iWiersz, iColumn + 1) = strWeekday
    
    If iWeekDay < 6 Then
    iWierszTeam = 2
    Do Until ThisWorkbook.Sheets("Team").Cells(iWierszTeam, "B") = iWeekDay
        iWierszTeam = iWierszTeam + 1
    Loop
    sMember = ThisWorkbook.Sheets("Team").Cells(iWierszTeam, "A")
    Else
    sMember = "Wolne"
    End If
    ThisWorkbook.Sheets("Kalendarz").Cells(iWiersz, iColumn + 2) = sMember
    
    
    dtData = DateAdd("d", 1, dtData)
    iWiersz = iWiersz + 1
Loop
     
dtData = DateAdd("m", -1, dtData)


End Sub

Sub Calendar()
Dim iCalendarMonth As Integer
Dim iCalendarColumn As Integer
Dim dtCalendarDate As Date
Dim iYear As Integer
Dim iQuestion As Integer


iYear = InputBox("Wprowadź rok:", "Grafik")
iQuestion = MsgBox("Czy chcesz wyświetlić grafik dla " & iYear & "?", vbYesNo + vbQuestion, "Grafik")
If iQuestion = vbYes Then
    GoTo continue
Else
    End
End If
continue:

If IsNumeric(iYear) = False Then End
iCalendarMonth = 1
iCalendarColumn = 1
dtCalendarDate = Format("01.01." & iYear, "dd.mm.yyyy")
Do While iCalendarMonth < 7 ' Untill iCalendarMonth=7
    Call oneMonth(iCalendarColumn, dtCalendarDate, iCalendarMonth)
    iCalendarColumn = iCalendarColumn + 3
    iCalendarMonth = iCalendarMonth + 1
    dtCalendarDate = DateAdd("m", 1, dtCalendarDate)
Loop
Call MsgBox("Odświeżanie ukończone:", vbOKOnly + vbInformation, "Grafik")
End Sub

#Dodawanie pracownika
Sub addEmployee()
ufPracownik.Show
End Sub
