VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolverParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D13:F13")
    OpenSolver.SetObjectiveFunctionCell Me.Range("H15")
    OpenSolver.AddConstraint Me.Range("H17:H18"), RelationLE, Me.Range("J17:J18")
    OpenSolver.AddConstraint Me.Range("H19"), RelationGE, Me.Range("J19")
    OpenSolver.AddConstraint Me.Range("H20"), RelationEQ, Me.Range("J20")
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetNonNegativity False
        
    Dim SolveResult As OpenSolverResult
    SolveResult = OpenSolver.RunOpenSolver(False, True)
    
    Select Case Solver
    ' Check CBC Options cause problem to solve correctly
    Case "CBC"
        If SolveResult = OpenSolverResult.Optimal And _
           Me.Range("A6").Value = True Then
            Test = Pass
        Else
            Test = Fail
        End If
    ' Check unbounded for all other solvers
    Case Else
        If SolveResult = OpenSolverResult.Unbounded Then
            Test = Pass
        Else
            Test = Fail
        End If
    End Select
End Function

