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
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D3:D35")
    OpenSolver.SetObjectiveFunctionCell Me.Range("I39")
    OpenSolver.AddConstraint Me.Range("J40"), RelationGE, Me.Range("L40")
    OpenSolver.AddConstraint Me.Range("J41:J42"), RelationLE, Me.Range("L41:L42")
    OpenSolver.AddConstraint Me.Range("J43"), RelationGE, Me.Range("L43")
    OpenSolver.AddConstraint Me.Range("J44"), RelationLE, Me.Range("L44")
    OpenSolver.AddConstraint Me.Range("D3:D35"), RelationINT
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetToleranceAsPercentage 1
        
    Test = RunTest(Me, Solver)
End Function
