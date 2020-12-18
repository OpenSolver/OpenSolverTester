VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("E5:E7"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("E10"), Me
    OpenSolver.SetObjectiveSense MaximiseObjective, Me
    OpenSolver.AddConstraint Me.Range("D10"), RelationLE, Me.Range("C10"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("E5:E7"), RelationBIN, sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
        
    Test = RunTest(Me, Solver)
End Function
