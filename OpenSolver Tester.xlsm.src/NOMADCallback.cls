VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NOMADCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function Test(Solver As String) As TestResult
    If Solver <> "NOMAD" Then
        Test = NA
        Exit Function
    End If

    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D6"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("D4"), Me
    OpenSolver.SetChosenSolver Solver, Me
    
    OpenSolver.AddConstraint Me.Range("D6"), RelationGE, RHSFormula:=-10, sheet:=Me
    OpenSolver.AddConstraint Me.Range("D6"), RelationLE, RHSFormula:=10, sheet:=Me
    
    ' With no callback, should fail
    OpenSolver.DeleteSolverParameters Solver, sheet:=Me
    Me.Range("D6").Value = 10
    Test = RunTest(Me, Solver)
    If Test <> Fail Then
        Test = Fail
        Exit Function
    End If
    
    ' Add callback so solve works
    OpenSolver.SetSolverParameters Solver, Me.Range("C10:D10"), sheet:=Me
    Me.Range("D6").Value = 10
    Test = RunTest(Me, Solver)
End Function


