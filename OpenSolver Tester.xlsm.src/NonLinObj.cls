VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLinObj"
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
    OpenSolver.AddConstraint Me.Range("H9"), RelationGE, Me.Range("J9")
    OpenSolver.AddConstraint Me.Range("H10"), RelationEQ, Me.Range("J10")
    OpenSolver.SetChosenSolver Solver
        
    Test = RunTest(Me, Solver)
End Function
