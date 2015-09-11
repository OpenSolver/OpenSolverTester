VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinIntOverlap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D4:G4"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("I6"), Me
    OpenSolver.AddConstraint Me.Range("I8:I9"), RelationLE, Me.Range("K8:K9"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I10"), RelationGE, Me.Range("K10"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I11"), RelationLE, Me.Range("K11"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("E4:F4"), RelationBIN, sheet:=Me
    OpenSolver.AddConstraint Me.Range("D4:F4"), RelationINT, sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
        
    Test = RunTest(Me, Solver)
End Function
