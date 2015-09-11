VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComplexLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    
    ' Hack for NOMAD range issues in non-english locales. See TestRunner for info
    OpenSolver.SetDecisionVariables Me.Range("D4:G4" & Separator & "Q3:R4" & Separator & "N3:N4"), Me
    
    OpenSolver.SetObjectiveFunctionCell Me.Range("I6"), Me
    OpenSolver.AddConstraint Me.Range("I8:I9"), RelationLE, Me.Range("K8:K9"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I10"), RelationGE, Me.Range("K10"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I11"), RelationLE, Me.Range("K11"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("E4:F4"), RelationBIN, sheet:=Me
    OpenSolver.AddConstraint Me.Range("D4:F4"), RelationINT, sheet:=Me
    OpenSolver.AddConstraint Me.Range("S3:S4"), RelationLE, Me.Range("U3:U4"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("Q5:R5"), RelationGE, Me.Range("Q7:R7"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
        
    Test = RunTest(Me, Solver)
End Function
