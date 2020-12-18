VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NamedRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D4:H4"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("I6"), Me
    OpenSolver.AddConstraint Me.Range("myNamedRange"), RelationLE, Me.Range("K8"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I9"), RelationGE, Me.Range("K9"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I10:I11"), RelationLE, Me.Range("K10:K11"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I12"), RelationEQ, Me.Range("K12"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    
    Test = RunTest(Me, Solver)
End Function
