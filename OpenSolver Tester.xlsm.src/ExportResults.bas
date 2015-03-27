Attribute VB_Name = "ExportResults"
Option Explicit

Public Sub ExportResults()

Dim ThisWB As Workbook, NewWB As Workbook
Dim ws As Worksheet
Dim SheetName As String
Dim i As Integer, j As Integer, MaxJ As Integer
Dim OfficeBitness As Integer

Application.ScreenUpdating = False

#If Win64 Then
    OfficeBitness = 64
#Else
    OfficeBitness = 32
#End If

Set ThisWB = ThisWorkbook
SheetName = "Export" & Format(DateTime.Now, "yyyy-MM-dd-hh-mm-ss")

MaxJ = 1
Do While ThisWB.Sheets("Results").Cells(2, MaxJ + 1).Value <> ""
    MaxJ = MaxJ + 1
Loop

Set NewWB = Workbooks.Add
Set ws = NewWB.Sheets(1)

i = 2
Do While ThisWB.Sheets("Results").Cells(i, 1).Value <> ""
    For j = 1 To MaxJ
        ws.Cells(i, j).Value = ThisWB.Sheets("Results").Cells(i, j).Value
    Next
    i = i + 1
Loop

ws.Cells(1, 1).NumberFormat = "@"
ws.Cells(1, 1) = Application.Version
ws.Cells(1, 2) = OfficeBitness
ws.Cells(1, 3) = OpenSolver.sOpenSolverVersion

Dim CSVName As String, CSVPath As String
CSVName = SheetName & ".csv"
CSVPath = ThisWB.Path + Application.PathSeparator & CSVName

ws.Activate
Application.DisplayAlerts = False
NewWB.SaveAs Filename:=CSVPath, FileFormat:=xlCSV

NewWB.Close
Application.DisplayAlerts = True

Dim ScriptPath As String
ScriptPath = ThisWorkbook.Path & Application.PathSeparator & "publish_results.py"
RunExternalCommand "python " & MakePathSafe(ScriptPath) & " " & MakePathSafe(CSVPath), "", Hide, True

Kill CSVPath
Application.ScreenUpdating = True


End Sub
