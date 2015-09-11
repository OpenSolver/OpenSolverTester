VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinIntLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D5:F5"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("H9"), Me
    OpenSolver.AddConstraint Me.Range("H11:H13"), RelationGE, Me.Range("J11:J13"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("D5:F5"), RelationGE, Me.Range("E7"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("F5"), RelationBIN, sheet:=Me
    OpenSolver.AddConstraint Me.Range("E5:F5"), RelationINT, sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    
    ' Reset value of binary to valid
    Me.Range("F5").Value = 0
    
    Test = RunTest(Me, Solver)
End Function


