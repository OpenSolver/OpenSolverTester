VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLin4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D11")
    OpenSolver.SetObjectiveFunctionCell Me.Range("D16")
    OpenSolver.SetChosenSolver Solver
    
    ' Avoid error on zero in cell
    Me.Range("D11").Value = 1
    
    If SolverLinearity(OpenSolver.CreateSolver(Solver)) = OpenSolver_SolverType.Linear Then
        Test = IIf(OpenSolver.RunOpenSolver(False, True) = ErrorOccurred, Pass, Fail)
    Else
        Test = RunTest(Me, Solver)
    End If
End Function


