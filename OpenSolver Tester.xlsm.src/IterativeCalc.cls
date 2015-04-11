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
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D5:F5")
    OpenSolver.SetObjectiveFunctionCell Me.Range("H7")
    OpenSolver.AddConstraint Me.Range("H9:H10"), RelationLE, Me.Range("J9:J10")
    OpenSolver.AddConstraint Me.Range("H11"), RelationGE, Me.Range("J11")
    OpenSolver.AddConstraint Me.Range("H12"), RelationEQ, Me.Range("J12")
    OpenSolver.SetChosenSolver Solver
    
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



