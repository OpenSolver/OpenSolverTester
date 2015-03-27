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
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D4:F4")
    OpenSolver.SetObjectiveFunctionCell Me.Range("G6")
    OpenSolver.AddConstraint Me.Range("G8"), RelationLE, Me.Range("I8")
    OpenSolver.AddConstraint Me.Range("D4:F4"), RelationINT
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetObjectiveSense MaximiseObjective
    
    Test = RunTest(Me, Solver)
End Function

