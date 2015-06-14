VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLinPruning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D5:E5"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("F6"), Me
    OpenSolver.AddConstraint Me.Range("F8:F9"), RelationLE, Me.Range("H8:H9"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("F10:F11"), RelationGE, Me.Range("H10:H11"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("F12:F18"), RelationLE, Me.Range("H12:H18"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    OpenSolver.SetObjectiveSense MaximiseObjective, Me
    
    ' Sometimes previous solution marked infeasible
    OpenSolver.SetPrecision 0.001, Me
    
    Test = RunTest(Me, Solver)
End Function

