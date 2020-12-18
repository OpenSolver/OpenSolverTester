VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IterativeCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D5:F5"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("H7"), Me
    OpenSolver.AddConstraint Me.Range("H9:H10"), RelationLE, Me.Range("J9:J10"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H11"), RelationGE, Me.Range("J11"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H12"), RelationEQ, Me.Range("J12"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    
    Application.Iteration = True
    
    ' Create a circular reference
    Me.Range("M9").Formula = "$N$9+1"
    Me.Range("N9").Formula = "$M$9+1"
    
    Test = RunTest(Me, Solver)
    
    'Delete circular reference
    Me.Range("M9").Clear
    Me.Range("N9").Clear
    
    Application.Iteration = False
End Function



