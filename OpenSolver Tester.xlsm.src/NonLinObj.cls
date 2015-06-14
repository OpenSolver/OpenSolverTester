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
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D3:F3"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("H5"), Me
    OpenSolver.AddConstraint Me.Range("H7:H8"), RelationLE, Me.Range("J7:J8"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H9"), RelationGE, Me.Range("J9"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H10"), RelationEQ, Me.Range("J10"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
        
    Test = RunTest(Me, Solver)
End Function
