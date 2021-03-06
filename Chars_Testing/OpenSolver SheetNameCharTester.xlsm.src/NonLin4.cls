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
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D11"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("D16"), Me
    OpenSolver.AddConstraint Me.Range("D11"), RelationLE, RHSFormula:=1000, sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    
    ' Avoid error on zero in cell
    Me.Range("D11").Value = 1
    
    If SolverLinearity(OpenSolver.CreateSolver(Solver)) = OpenSolver_SolverType.Linear Then
        Test = IIf(OpenSolver.RunOpenSolver(False, True) = ErrorOccurred, Pass, Fail)
    Else
        Test = RunTest(Me, Solver)
    End If
End Function


