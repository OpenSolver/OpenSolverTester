VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D3:D35"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("I39"), Me
    OpenSolver.AddConstraint Me.Range("J40"), RelationGE, Me.Range("L40"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("J41:J42"), RelationLE, Me.Range("L41:L42"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("J43"), RelationGE, Me.Range("L43"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("J44"), RelationLE, Me.Range("L44"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("D3:D35"), RelationINT, sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    OpenSolver.SetToleranceAsPercentage 1, Me
        
    Test = RunTest(Me, Solver)
End Function
