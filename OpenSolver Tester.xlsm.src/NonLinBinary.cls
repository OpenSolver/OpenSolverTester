VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLinBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("C4:D4")
    OpenSolver.SetObjectiveFunctionCell Me.Range("C6")
    OpenSolver.AddConstraint Me.Range("E8:E9"), RelationGE, Me.Range("G8:G9")
    OpenSolver.AddConstraint Me.Range("C4:D4"), RelationBIN
    OpenSolver.SetChosenSolver Solver
        
    Test = RunTest(Me, Solver)
End Function


