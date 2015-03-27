VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndirectLBs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D5:F5")
    OpenSolver.SetObjectiveFunctionCell Me.Range("H8")
    OpenSolver.AddConstraint Me.Range("H10:H11"), RelationLE, Me.Range("J10:J11")
    OpenSolver.AddConstraint Me.Range("H12"), RelationGE, Me.Range("J12")
    OpenSolver.AddConstraint Me.Range("H13"), RelationEQ, Me.Range("J13")
    OpenSolver.AddConstraint Me.Range("D5:E5"), RelationGE, Me.Range("D6:E6")
    OpenSolver.AddConstraint Me.Range("D5"), RelationLE, RHSFormula:="2"
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetMaxIterations 2000
    
    Test = RunTest(Me, Solver)
End Function


