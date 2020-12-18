VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLin5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("F2:I2"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("J7"), Me
    OpenSolver.SetChosenSolver Solver, Me
    
    ' Reset decision vars so non-linear solvers avoid an evaluating sqrt(0)
    Me.Range("F2").Value = 1
    Me.Range("G2").Value = 2
    Me.Range("H2").Value = 3
    Me.Range("I2").Value = 4
    
    Test = RunTest(Me, Solver)
End Function

