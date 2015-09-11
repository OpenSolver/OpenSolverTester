VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLinMinMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D4:F4"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("G6"), Me
    OpenSolver.AddConstraint Me.Range("G8"), RelationLE, Me.Range("I8"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("D4:F4"), RelationINT, sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    OpenSolver.SetObjectiveSense MaximiseObjective, Me
    
    Test = RunTest(Me, Solver)
End Function

