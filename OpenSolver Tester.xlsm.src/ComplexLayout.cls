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
    Me.Activate
    OpenSolver.ResetModel
    
    ' Hack for NOMAD range issues in non-english locales. See TestRunner for info
    OpenSolver.SetDecisionVariables Me.Range("D4:G4" & Separator & "Q3:R4" & Separator & "N3:N4")
    
    OpenSolver.SetObjectiveFunctionCell Me.Range("I6")
    OpenSolver.AddConstraint Me.Range("I8:I9"), RelationLE, Me.Range("K8:K9")
    OpenSolver.AddConstraint Me.Range("I10"), RelationGE, Me.Range("K10")
    OpenSolver.AddConstraint Me.Range("I11"), RelationLE, Me.Range("K11")
    OpenSolver.AddConstraint Me.Range("E4:F4"), RelationBIN
    OpenSolver.AddConstraint Me.Range("D4:F4"), RelationINT
    OpenSolver.AddConstraint Me.Range("S3:S4"), RelationLE, Me.Range("U3:U4")
    OpenSolver.AddConstraint Me.Range("Q5:R5"), RelationGE, Me.Range("Q7:R7")
    OpenSolver.SetChosenSolver Solver
        
    Test = RunTest(Me, Solver)
End Function
