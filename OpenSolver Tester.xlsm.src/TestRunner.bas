Attribute VB_Name = "TestRunner"
Option Explicit

' Test Runner for testing sheet.
' All tests should have the testing column as Column A on the sheet. These cells control the
' parameters for testing each sheet. There are two types of test:
'     - Normal, which is tested by the NormalTest.NormalTest function
'     - Custom, which is tested by a sheet-specific function called CustomTest.<SheetName>
'
' Model Type determines which solvers are tested on the problem - linear solvers are only tested on
' linear problems, while non-linear solvers are tested on all problems.
'
' Parameters specified in the testing column for Normal test:
'     - CorrectResult, takes value TRUE if solver output on sheet is as expected.
'     - ExpectedResult, specifies the OpenSolver code that should be returned.

Sub RunAllTests()
    Dim ListCount As Integer
    
    ' Get linear solvers
    Dim LinearSolvers As Collection
    Set LinearSolvers = New Collection
    With Sheet53.ListBoxLinear
        For ListCount = 0 To .ListCount - 1
            If .Selected(ListCount) Then
                LinearSolvers.Add .List(ListCount)
            End If
        Next ListCount
    End With

    ' Get non-linear solvers
    Dim NonLinearSolvers As Collection
    Set NonLinearSolvers = New Collection
    With Sheet53.ListBoxNonLinear
        For ListCount = 0 To .ListCount - 1
            If .Selected(ListCount) Then
                NonLinearSolvers.Add .List(ListCount)
            End If
        Next ListCount
    End With
    
    ' Set up results sheet
    Dim Solver As Variant, ProblemType As String, Result As Variant
    Dim i As Integer, j As Integer
    i = 0
    j = 0
    Sheets("Results").Cells.ClearContents
    SetResultCell i, j, "Test"
    
    For Each Solver In LinearSolvers
        j = j + 1
        SetResultCell i, j, Solver
    Next Solver
    For Each Solver In NonLinearSolvers
        j = j + 1
        SetResultCell i, j, Solver
    Next Solver
    
    
    Dim SheetName As String, listIndex As Integer
    With Sheet53.ListBoxTest
        For listIndex = 0 To .ListCount - 1
            ' Exit if not selected
            If .Selected(listIndex) = False Then
                GoTo NextSheet
            End If
            
            ' Add test to results sheet
            SheetName = .List(listIndex)
            i = i + 1
            j = 0
            SetResultCell i, j, SheetName
            Sheets(SheetName).Activate
            
            ' Read problem type and test the appropriate solvers for each test
            ProblemType = Sheets(Sheet53.ListBoxTest.List(listIndex)).Cells(4, 1)
            For Each Solver In LinearSolvers
                j = j + 1
                If ProblemType = "Linear" Then
                    SetResultCell i, j, FormatResult(RunTest(Sheets(SheetName), Solver, True))
                Else
                    SetResultCell i, j, FormatResult(RunNonLinearityTest(Sheets(SheetName), Solver))
                End If
            Next Solver
            
            For Each Solver In NonLinearSolvers
                j = j + 1
                SetResultCell i, j, FormatResult(RunTest(Sheets(SheetName), Solver, True))
            Next Solver
        
NextSheet:
        Next listIndex
    End With
    
    Sheets("Results").Activate
End Sub

Function RunTest(Sheet As Worksheet, Solver As Variant, ReturnValidation As Boolean)
' Runs a test problem for a single solver
    Dim VBComp As Variant, SolveResult As Integer
    Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & Solver)
    If Sheet.Cells(2, 1) = "Normal" Then
        If ReturnValidation = True Then
            RunTest = NormalTest.NormalTest(Sheet)
        Else
            RunTest = NormalTest.NormalTestWithoutReturnValidation(Sheet)
        End If
    Else
        RunTest = Application.Run(Sheet.Name, Sheet, Solver)
    End If
    
    If TestShouldFail(Sheet.Name, CStr(Solver)) Then
        RunTest = RunTest + 10
    End If
End Function

Function RunNonLinearityTest(Sheet As Worksheet, Solver As Variant)
    Dim VBComp As Variant, SolveResult As Integer
    Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & Solver)
    If Sheet.Cells(2, 1) = "Normal" Then
        RunNonLinearityTest = NormalTest.NonLinearityTest(Sheet)
    Else
        RunNonLinearityTest = Application.Run(Sheet.Name, Sheet, Solver)
    End If
    
    If TestShouldFail(Sheet.Name, CStr(Solver)) Then
        RunNonLinearityTest = RunNonLinearityTest + 10
    End If
End Function

Sub SetResultCell(i As Integer, j As Integer, Result As Variant)
    Sheets("Results").Cells(1 + i, 2 + j).Value = Result
End Sub

Function FormatResult(Result As Variant)
    Select Case Result
    Case 1
        FormatResult = "PASS"
    Case 11 ' a passing test on the fail whitelist
        FormatResult = "PASS*"
    Case 0
        FormatResult = "FAIL"
    Case 10 ' a failing test on the fail whitelist
        FormatResult = "FAIL*"
    Case -1
        FormatResult = "N/A"
    Case 9 ' an N/A test on the fail whitelist
        FormatResult = "N/A*"
    Case Else
        FormatResult = Result
    End Select
End Function

Sub RefreshTestListBox()
    With Sheet53.ListBoxTest
        .Clear
    
        ' Loop through worksheets looking for sheets with the testing pane present.
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ' Exit if not a testing sheet
            If ws.Cells(2, 1).Value = "Normal" Or ws.Cells(2, 1).Value = "Custom" Then
                .AddItem ws.Name
            End If
        Next ws
        
        ' Fix scrolling issue: https://stackoverflow.com/questions/5859459
        .IntegralHeight = False
        .Height = 250
        .Width = Sheet53.Range("A17").Width - 20
        .IntegralHeight = True
        .MultiSelect = fmMultiSelectMulti
        
        .Enabled = True
    End With
    ' Decheck select all box
    Sheet53.Shapes("CheckBoxTest").OLEFormat.Object.Value = -4146
End Sub

Sub SelectAll()
    With Sheet53.ListBoxTest
        If Sheet53.Shapes("CheckBoxTest").OLEFormat.Object.Value = 1 Then
            Dim i As Integer
            For i = 0 To .ListCount - 1
                .Selected(i) = True
            Next i
            .Enabled = False
        Else
            .Enabled = True
        End If
    End With
End Sub

