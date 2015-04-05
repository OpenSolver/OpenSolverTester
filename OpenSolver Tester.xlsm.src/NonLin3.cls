VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLin3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("S10:S11")
    OpenSolver.SetObjectiveFunctionCell Me.Range("S13")
    OpenSolver.AddConstraint Me.Range("S10:S11"), RelationGE, Me.Range("Q10:Q11")
    OpenSolver.AddConstraint Me.Range("S10:S11"), RelationLE, Me.Range("U10:U11")
    OpenSolver.SetChosenSolver Solver
    
    Test = RunTest(Me, Solver)
End Function


