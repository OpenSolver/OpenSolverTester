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
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D6:F6"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("H10"), Me
    OpenSolver.AddConstraint Me.Range("H12:H14"), RelationGE, Me.Range("J12:J14"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("E6:H6"), RelationGE, Me.Range("E8:H8"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    Test = RunTest(Me, Solver)
End Function


