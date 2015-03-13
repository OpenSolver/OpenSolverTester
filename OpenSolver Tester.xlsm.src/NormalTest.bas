Attribute VB_Name = "NormalTest"
Function NormalTest(Sheet As Worksheet, Optional SolveRelaxation As Boolean = False)
' Checks OpenSolver code and sheet output are as expected.
    Dim SolveResult As Integer
    SolveResult = OpenSolver.RunOpenSolver(SolveRelaxation, True)
    Application.Calculate
    If SolveResult = Sheet.Range("A9").Value And _
       Sheet.Range("A6").Value = True Then
        NormalTest = 1 ' PASS
    Else
        NormalTest = 0 ' FAIL
    End If
End Function

Function NormalTestWithoutReturnValidation(Sheet As Worksheet)
' Checks sheet output is as expected. Used for testing a solver that doesn't have return codes hooked up
    Dim SolveResult As Integer
    SolveResult = OpenSolver.RunOpenSolver(False, True)
    Application.Calculate
    If Sheet.Range("A6").Value = True Then
        NormalTestWithoutReturnValidation = 1 ' PASS
    Else
        NormalTestWithoutReturnValidation = 0 ' FAIL
    End If
End Function

Function NonLinearityTest(Sheet As Worksheet)
' Checks that OpenSolver outputs a 'NotLinear' return code
    Dim SolveResult As Integer
    SolveResult = OpenSolver.RunOpenSolver(False, True, 10)
    If SolveResult = OpenSolverResult.NotLinear Then
        NonLinearityTest = 1 ' PASS
    Else
        NonLinearityTest = 0 ' FAIL
    End If
End Function
