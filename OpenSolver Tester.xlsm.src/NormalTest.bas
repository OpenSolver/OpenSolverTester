Attribute VB_Name = "NormalTest"
Function NormalTest(sheet As Worksheet, Optional SolveRelaxation As Boolean = False) As TestResult
' Checks OpenSolver code and sheet output are as expected.
    Dim SolveResult As OpenSolverResult
    SolveResult = OpenSolver.RunOpenSolver(SolveRelaxation, True, sheet:=sheet)
    Application.Calculate
    If SolveResult = sheet.Range("A9").Value And _
       sheet.Range("A6").Value = True Then
        NormalTest = Pass
    Else
        NormalTest = Fail
    End If
End Function

Function NonLinearityTest(sheet As Worksheet) As TestResult
' Checks that OpenSolver outputs a 'NotLinear' return code
    Dim SolveResult As Integer
    SolveResult = OpenSolver.RunOpenSolver(False, True, 10, sheet)
    If SolveResult = OpenSolverResult.NotLinear Then
        NonLinearityTest = Pass
    Else
        NonLinearityTest = Fail
    End If
End Function
