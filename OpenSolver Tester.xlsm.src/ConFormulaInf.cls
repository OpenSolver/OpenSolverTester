VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConFormulaInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D5:G5,Q4:R5,N4:N5")
    OpenSolver.SetObjectiveFunctionCell Me.Range("I7")
    OpenSolver.AddConstraint Me.Range("I9:I10"), RelationLE, Me.Range("K9:K10")
    OpenSolver.AddConstraint Me.Range("K11"), RelationLE, Me.Range("N11")
    OpenSolver.AddConstraint Me.Range("I12"), RelationLE, Me.Range("K12")
    OpenSolver.AddConstraint Me.Range("E5:F5"), RelationBIN
    OpenSolver.AddConstraint Me.Range("D5:F5"), RelationINT
    OpenSolver.AddConstraint Me.Range("S4:S5"), RelationLE, Me.Range("U4:U5")
    OpenSolver.AddConstraint Me.Range("Q6:R6"), RelationGE, Me.Range("Q8:R8")
    OpenSolver.AddConstraint Me.Range("I10"), RelationGE, RHSFormula:="=$I$11/4+$B$3"
    OpenSolver.AddConstraint Me.Range("I14"), RelationLE, RHSFormula:="=5*$I$15-1"
    OpenSolver.SetChosenSolver Solver
        
    Test = RunTest(Me, Solver)
End Function


