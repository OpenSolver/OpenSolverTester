VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuickSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D4:F4")
    OpenSolver.SetObjectiveFunctionCell Me.Range("H6")
    OpenSolver.AddConstraint Me.Range("H8"), RelationLE, Me.Range("J8")
    OpenSolver.AddConstraint Me.Range("H9:H10"), RelationGE, Me.Range("J9:J10")
    OpenSolver.SetChosenSolver Solver
    
    
    If SolverType(Solver) <> OpenSolver_SolverType.Linear Then
        Test = NA
        Exit Function
    End If
    
    OpenSolver.SetQuickSolveParameters Me.Range("O7:P7")
    OpenSolver.InitializeQuickSolve
    
    ' Test first set of values
    Dim SolveResult As OpenSolverResult
    Dim CorrectResult1 As Boolean
    Me.Range("Scale").Value = -2
    Me.Range("Offset").Value = 4
    SolveResult = OpenSolver.RunQuickSolve(True)
    CorrectResult1 = Me.Range("H16").Value And SolveResult = OpenSolverResult.Optimal
    
    ' Test second set of values
    Dim CorrectResult2 As Boolean
    Me.Range("Scale").Value = 2.5
    Me.Range("Offset").Value = -50
    SolveResult = OpenSolver.RunQuickSolve(True)
    CorrectResult2 = Me.Range("H20").Value And SolveResult = OpenSolverResult.Optimal
    
    If CorrectResult1 And CorrectResult2 Then
        Test = Pass
    Else
        Test = Fail
    End If
End Function

