VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeasConstConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D4:F4"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("J6"), Me
    OpenSolver.AddConstraint Me.Range("J8:J9"), RelationLE, Me.Range("L8:L9"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("J10"), RelationGE, Me.Range("L10"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("J11"), RelationGE, Me.Range("L11"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    
    Test = RunTest(Me, Solver)
End Function

