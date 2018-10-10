Attribute VB_Name = "mMethods"
Option Explicit

Sub screenUpdating()
'Disable screenUpdating for when the macro is running and enable before the end sub
'This stops the screen from flickering and reduces the macro running time

    Application.screenUpdating = False
    Application.screenUpdating = True
    
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    
End Sub
Sub variableAssignment()

    Dim userValue As Long
    Dim userName As String
    Dim userDate As Date
    Dim wSheet As Worksheet
    
    userValue = 10
    userName = "Joe Bloggs"
    userDate = Date '"Date" returns the current date
    
    Set wSheet = ThisWorkbook.Worksheets("TestA")
    'or
    Set wSheet = ActiveSheet
    
End Sub

Sub copyPaste()
    
    'Change "xlValues" to another given argument to alter the paste type (formulas/formats etc)
    ThisWorkbook.Worksheets("Macros").Range("I10:I12").Copy
    ThisWorkbook.Worksheets("Macros").Range("J10").PasteSpecial xlValues
    
    'Or (Only works for values...)
    ThisWorkbook.Worksheets("Macros").Range("J10:J12").Value = ThisWorkbook.Worksheets("Macros").Range("I10:I12").Value
    
    Application.CutCopyMode = False
    
End Sub

Sub hideAndUnhideWorksheet()
    'xlSheetVeryHidden to hide sheet so it does not show in the list of hidden sheets
    
    ThisWorkbook.Worksheets("Test1").Visible = xlSheetHidden
    
    ThisWorkbook.Worksheets("Test1").Visible = xlSheetVisible
    
End Sub

Sub hideAndUnhideAllSheets()
' To hide all worksheets change 'xlSheetVisible' to 'xlSheetHidden' or 'xlSheetVeryHidden'
    Dim ws As Worksheet
    
    'Hides all sheets except active sheet and one worksheet must always be presest
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.Name = ActiveSheet.Name Then
            ws.Visible = xlSheetHidden
        End If
    Next ws
    
    'unhides all sheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws

End Sub

Sub ifTest()

    Dim userValue As Long
    userValue = ThisWorkbook.Worksheets("TestA").Range("B1").Value
    
    If userValue < 10 Then
        MsgBox "<10", vbOKOnly
    Else
        MsgBox ">=10", vbOKOnly
    End If
    
End Sub

Sub ifElseTest()

    Dim userValue As Long
    userValue = ThisWorkbook.Worksheets("TestA").Range("B1")
    
    If userValue < 10 Then
        MsgBox "<10", vbOKOnly
    ElseIf userValue = 10 Then
        MsgBox "=10", vbOKOnly
    Else
        MsgBox ">10", vbOKOnly
    End If

End Sub

Sub forLoop()
'Iterates through each value within a set of values

    Dim userAge As Long
    Dim i As Long
    i = 1
    
    Dim bottomDataRow As Long
    bottomDataRow = ThisWorkbook.Worksheets("2").Cells(Rows.Count, 6).End(xlUp).Row
    
    For i = 4 To bottomDataRow
        
        userAge = ThisWorkbook.Worksheets("2").Range("G" & i).Value
        
        If userAge > 20 Then
            ThisWorkbook.Worksheets("2").Range("H" & i).Value = "True"
        Else
            ThisWorkbook.Worksheets("2").Range("H" & i).Value = "False"
        End If
        
    Next i

End Sub

Sub forEachLoop()
'Iterates through each object within a group of objects

    Dim wSheet As Worksheet
    
    For Each wSheet In ThisWorkbook.Worksheets
        wSheet.Range("A100").Value = 100
    Next wSheet
    
End Sub

Sub loopsNotCovered()

'Do until    Loop until event is true
'Do while    Loop while event is true

End Sub
Sub findBottomRow()

    'Amend worksheet name and column number
    'This will retrieve the last used row containing data in a given column (Here it's column 8 in the "Macros" worksheet
    Dim bottomDataRow As Long
    bottomDataRow = ThisWorkbook.Worksheets("Macros").Cells(Rows.Count, 8).End(xlUp).Row
    
End Sub

    
