VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextRHS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D4:F4"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("H6"), Me
    OpenSolver.AddConstraint Me.Range("H8:H9"), RelationLE, Me.Range("J8:J9"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H10"), RelationGE, Me.Range("J10"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H11"), RelationEQ, Me.Range("J11"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
        
    Test = RunTest(Me, Solver)
End Function
