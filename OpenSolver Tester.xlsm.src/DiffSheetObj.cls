VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffSheetObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D2:F2"), Me
    OpenSolver.SetObjectiveFunctionCell SeekObj.Range("G17"), Me
    OpenSolver.AddConstraint Me.Range("G6"), RelationEQ, Me.Range("I6"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("G7"), RelationGE, Me.Range("I7"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("G8:G9"), RelationLE, Me.Range("I8:I9"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    
    Test = RunTest(Me, Solver)
End Function


