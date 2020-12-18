VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLinConstraints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D5:H5"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("I6"), Me
    OpenSolver.AddConstraint Me.Range("I7:I9"), RelationLE, Me.Range("K7:K9"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    
    Test = RunTest(Me, Solver)
End Function


