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
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("F2:I2")
    OpenSolver.SetObjectiveFunctionCell Me.Range("J7")
    OpenSolver.SetChosenSolver Solver
    
    ' Reset decision vars so non-linear solvers avoid an evaluating sqrt(0)
    Me.Range("F2").Value = 1
    Me.Range("G2").Value = 2
    Me.Range("H2").Value = 3
    Me.Range("I2").Value = 4
    
    Test = RunTest(Me, Solver)
End Function

