VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLin6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D7:D8"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("F15"), Me
    OpenSolver.SetChosenSolver Solver, Me
    
    ' Starting point 1
    Me.Range("D7").Value = -30
    Me.Range("D8").Value = 50
    
    Test = RunTest(Me, Solver)
    If Test <> Pass Then Exit Function
    
    ' Starting point 2
    Me.Range("D7").Value = 30
    Me.Range("D8").Value = -50
    
    Test = RunTest(Me, Solver)
End Function

