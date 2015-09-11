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
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D5:F5"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("H8"), Me
    OpenSolver.AddConstraint Me.Range("H10:H11"), RelationLE, Me.Range("J10:J11"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H12"), RelationGE, Me.Range("J12"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H13"), RelationEQ, Me.Range("J13"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("D5:E5"), RelationGE, Me.Range("D6:E6"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("D5"), RelationLE, RHSFormula:="2", sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    
    Test = RunTest(Me, Solver)
End Function


