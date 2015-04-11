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
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D5:E5")
    OpenSolver.SetObjectiveFunctionCell Me.Range("F6")
    OpenSolver.AddConstraint Me.Range("F8:F9"), RelationLE, Me.Range("H8:H9")
    OpenSolver.AddConstraint Me.Range("F10:F11"), RelationGE, Me.Range("H10:H11")
    OpenSolver.AddConstraint Me.Range("F12:F18"), RelationLE, Me.Range("H12:H18")
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetObjectiveSense MaximiseObjective
    
    ' Sometimes previous solution marked infeasible
    OpenSolver.SetPrecision 0.001
    
    Test = RunTest(Me, Solver)
End Function

