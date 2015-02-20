Attribute VB_Name = "CustomTest"
Function Test15e_Relax(Sheet As Worksheet, Solver As Variant)
    'Test15e_Relax = NormalTest.NormalTest(Sheet, True)
    SolveResult = RunOpenSolver(False, True)
    If SolveResult <> OpenSolverResult.ErrorOccurred Then
        Test15e_Relax = 0 ' FAIL
        Exit Function
    End If
    
    SolveResult = RunOpenSolver(True, True)
    If SolveResult = OpenSolverResult.Optimal And _
        Sheet.Range("A6").Value = True Then
        Test15e_Relax = 1 ' PASS
    Else
        Test15e_Relax = 0 ' FAIL
    End If

End Function

Function Test24_IterativeCalc(Sheet As Worksheet, Solver As Variant)
    Application.Iteration = True
    Test24_IterativeCalc = NormalTest.NormalTest(Sheet)
    Application.Iteration = False
End Function

Function Test27_Parameters(Sheet As Worksheet, Solver As Variant)
    If SolverType(CStr(Solver)) <> OpenSolver_SolverType.Linear Then
        Test27_Parameters = -1 ' N/A
        Exit Function
    End If
    
    InitializeQuickSolve
    
    ' Test first set of values
    Dim SolveResult As Integer
    Dim CorrectResult1 As Boolean
    Sheet.Range("Scale").Value = -2
    Sheet.Range("Offset").Value = 4
    SolveResult = RunQuickSolve(True)
    CorrectResult1 = Sheet.Range("H16").Value And SolveResult = OpenSolverResult.Optimal
    
    ' Test second set of values
    Dim CorrectResult2 As Boolean
    Sheet.Range("Scale").Value = 2.5
    Sheet.Range("Offset").Value = -50
    SolveResult = RunQuickSolve(True)
    CorrectResult2 = Sheet.Range("H20").Value And SolveResult = OpenSolverResult.Optimal
    
    If CorrectResult1 And CorrectResult2 Then
        Test27_Parameters = 1 ' PASS
    Else
        Test27_Parameters = 0 ' FAIL
    End If
End Function

Function Test28_CBCOptions(Sheet As Worksheet, Solver As Variant)
    Dim SolveResult As Integer
    SolveResult = RunOpenSolver(False, True)
    
    ' Check CBC Options cause problem to solve correctly
    If Solver = "CBC" Then
        If SolveResult = OpenSolverResult.Optimal And _
           Sheet.Range("A6").Value = True Then
            Test28_CBCOptions = 1 ' PASS
        Else
            Test28_CBCOptions = 0 ' FAIL
        End If
    ' Check unbounded for all other solvers
    Else
        If SolveResult = OpenSolverResult.Unbounded Then
            Test28_CBCOptions = 1 ' PASS
        Else
            Test28_CBCOptions = 0 ' FAIL
        End If
    End If
End Function

Function Test40(Sheet As Worksheet, Solver As Variant)
    Dim SolveResult As Integer
    ' Reset decision var to 1 so non-linear solvers avoid an error
    Sheet.Range("D11").Value = 1
    SolveResult = RunOpenSolver(False, True)
    If SolverType(CStr(Solver)) = OpenSolver_SolverType.Linear Then
        ' Check linear solvers get an error in the objective cell
        If SolveResult = OpenSolverResult.ErrorOccurred Then
            Test40 = 1 ' PASS
        Else
            Test40 = 0 ' FAIL
        End If
    Else
        ' Check non-linear solvers are correct
        If SolveResult = Sheet.Range("A9").Value And _
           Sheet.Range("A6").Value = True Then
            Test40 = 1 ' PASS
        Else
            Test40 = 0 ' FAIL
        End If
    End If
End Function

Function Test41(Sheet As Worksheet, Solver As Variant)
    Dim SolveResult As Integer
    ' Reset decision vars so non-linear solvers avoid an evaluating sqrt(0)
    Sheet.Range("F2").Value = 1
    Sheet.Range("G2").Value = 2
    Sheet.Range("H2").Value = 3
    Sheet.Range("I2").Value = 4
    SolveResult = RunOpenSolver(False, True, 10)
    If SolverType(CStr(Solver)) = OpenSolver_SolverType.Linear Then
        ' Check linear solvers return NotLinear
        If SolveResult = OpenSolverResult.NotLinear Then
            Test41 = 1 ' PASS
        Else
            Test41 = 0 ' FAIL
        End If
    Else
        ' Check non-linear solvers are correct
        If SolveResult = Sheet.Range("A9").Value And _
           Sheet.Range("A6").Value = True Then
            Test41 = 1 ' PASS
        Else
            Test41 = 0 ' FAIL
        End If
    End If
End Function
