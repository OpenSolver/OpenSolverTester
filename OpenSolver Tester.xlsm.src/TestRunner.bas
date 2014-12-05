Attribute VB_Name = "TestRunner"
Option Explicit

Sub LaunchForm()
   TestSelector.Show
End Sub

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
    With TestSelector.lstLinearSolvers
        For ListCount = 0 To .ListCount - 1
            If .Selected(ListCount) Then
                LinearSolvers.Add .List(ListCount)
            End If
        Next ListCount
    End With

    ' Get non-linear solvers
    Dim NonLinearSolvers As Collection
    Set NonLinearSolvers = New Collection
    With TestSelector.lstNonLinearSolvers
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
    Sheets("Results").Cells(1, 3).Value = "Tests marked with * are expected to fail. Look at the test whitelist module for info on why they should fail"
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
    With TestSelector.lstTests
        For listIndex = 0 To .ListCount - 1
            ' Exit if not selected
            If .Selected(listIndex) = False Then
                GoTo NextSheet
            End If
            
            ' Add test to results sheet
            SheetName = .List(listIndex)
            i = i + 1
            j = 0
            SetResultCell i, j, "=HYPERLINK(""[OpenSolver Tester.xlsm]" & SheetName & "!A1"", """ & SheetName & """)"
            Sheets(SheetName).Activate
            
            ' Read problem type and test the appropriate solvers for each test
            ProblemType = Sheets(.List(listIndex)).Cells(4, 1)
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
    OpenSolver.SetNameOnSheet "OpenSolver_ChosenSolver", "=" & Solver
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
    Sheets("Results").Cells(2 + i, 1 + j).Value = Result
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

