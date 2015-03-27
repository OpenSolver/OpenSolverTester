VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OverlapRangeLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D6:F6")
    OpenSolver.SetObjectiveFunctionCell Me.Range("H10")
    OpenSolver.AddConstraint Me.Range("H12:H14"), RelationGE, Me.Range("J12:J14")
    OpenSolver.AddConstraint Me.Range("E6:H6"), RelationGE, Me.Range("E8:H8")
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetMaxIterations 2000
    
    Test = RunTest(Me, Solver)
End Function


