VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectLBs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D4:F4")
    OpenSolver.SetObjectiveFunctionCell Me.Range("H8")
    OpenSolver.AddConstraint Me.Range("H10:H12"), RelationGE, Me.Range("J10:J12")
    OpenSolver.AddConstraint Me.Range("D4:F4"), RelationGE, Me.Range("D6:F6")
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetMaxIterations 2000
    
    Test = RunTest(Me, Solver)
End Function



