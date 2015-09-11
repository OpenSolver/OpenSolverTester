VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeletedConRef"
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
    OpenSolver.AddConstraint Me.Range("H12"), RelationLE, Me.Range("J12"), sheet:=Me
    
    ' Add constraint and delete ref afterwards
    OpenSolver.AddConstraint Me.Range("K1"), RelationLE, Me.Range("K2"), sheet:=Me
    Me.Range("K2").Delete
    
    OpenSolver.SetChosenSolver Solver, Me
        
    Test = RunTest(Me, Solver)
End Function



