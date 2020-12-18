VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    
    ' Hack for NOMAD range issues in non-english locales. See TestRunner for info
    OpenSolver.SetDecisionVariables Me.Range("D5:G5" & Separator & "Q4:R5" & Separator & "N4:N5"), Me
    
    OpenSolver.SetObjectiveFunctionCell Me.Range("I7"), Me
    OpenSolver.AddConstraint Me.Range("I9:I10"), RelationLE, Me.Range("K9:K10"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I11"), RelationGE, Me.Range("K11"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I12"), RelationLE, Me.Range("K12"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("E5:F5"), RelationBIN, sheet:=Me
    OpenSolver.AddConstraint Me.Range("D5:F5"), RelationINT, sheet:=Me
    OpenSolver.AddConstraint Me.Range("S4:S5"), RelationLE, Me.Range("U4:U5"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("Q6:R6"), RelationGE, Me.Range("Q8:R8"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("I10"), RelationGE, RHSFormula:="=$I$11/4+$B$3", sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
        
    Test = RunTest(Me, Solver)
End Function

