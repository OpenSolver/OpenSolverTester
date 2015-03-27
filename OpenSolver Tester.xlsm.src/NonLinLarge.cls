VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLinLarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D3:F3")
    OpenSolver.SetObjectiveFunctionCell Me.Range("H5")
    OpenSolver.AddConstraint Me.Range("H7:H8"), RelationLE, Me.Range("J7:J8")
    OpenSolver.AddConstraint Me.Range("H9:H72"), RelationGE, Me.Range("J9:J72")
    OpenSolver.AddConstraint Me.Range("H73"), RelationEQ, Me.Range("J73")
    OpenSolver.AddConstraint Me.Range("H74"), RelationLE, Me.Range("J74")
    OpenSolver.AddConstraint Me.Range("H75:H138"), RelationGE, Me.Range("J75:J138")
    OpenSolver.AddConstraint Me.Range("H139"), RelationEQ, Me.Range("J139")
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetMaxIterations 2000
        
    Test = RunTest(Me, Solver)
End Function

