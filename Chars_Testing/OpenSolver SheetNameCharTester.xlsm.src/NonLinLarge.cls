VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonLinLarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D3:F3"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("H5"), Me
    
    'Dim i As Long
    'For i = 1 To 100
        OpenSolver.AddConstraint Me.Range("H7:H8"), RelationLE, Me.Range("J7:J8"), sheet:=Me
        OpenSolver.AddConstraint Me.Range("H9:H72"), RelationGE, Me.Range("J9:J72"), sheet:=Me
        OpenSolver.AddConstraint Me.Range("H73"), RelationEQ, Me.Range("J73"), sheet:=Me
        OpenSolver.AddConstraint Me.Range("H74"), RelationLE, Me.Range("J74"), sheet:=Me
        OpenSolver.AddConstraint Me.Range("H75:H138"), RelationGE, Me.Range("J75:J138"), sheet:=Me
        OpenSolver.AddConstraint Me.Range("H139"), RelationEQ, Me.Range("J139"), sheet:=Me
    'Next i
    OpenSolver.SetChosenSolver Solver, Me
    
    Dim starttime As Single
    starttime = Timer()
    Test = RunTest(Me, Solver)
    Debug.Print Timer() - starttime
End Function
