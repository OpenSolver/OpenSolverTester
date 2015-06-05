VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeekObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D2:F2")
    OpenSolver.SetObjectiveFunctionCell Me.Range("G4")
    OpenSolver.AddConstraint Me.Range("G6"), RelationEQ, Me.Range("I6")
    OpenSolver.AddConstraint Me.Range("G7"), RelationGE, Me.Range("I7")
    OpenSolver.AddConstraint Me.Range("G8:G9"), RelationLE, Me.Range("I8:I9")
    OpenSolver.SetChosenSolver Solver
    
    OpenSolver.SetObjectiveSense TargetObjective
    OpenSolver.SetObjectiveTargetValue 50
    OpenSolver.SetDualsOnSheet True
    OpenSolver.SetUpdateSensitivity True
    
    Test = RunTest(Me, Solver)
    If SensitivityAnalysisAvailable(CreateSolver(Solver)) Then
        Dim SensitivitySheet As Worksheet
        On Error GoTo ErrorHandler
        Set SensitivitySheet = Sheets(Me.Name & " Sensitivity")
        On Error GoTo 0
        
        If SensitivitySheet.Range("E7").Value <> 0 Or _
           SensitivitySheet.Range("E8").Value <> 0 Or _
           SensitivitySheet.Range("E9").Value <> 0 Or _
           SensitivitySheet.Range("E13").Value <> 0 Or _
           SensitivitySheet.Range("E14").Value <> 0 Or _
           SensitivitySheet.Range("E15").Value <> 0 Or _
           SensitivitySheet.Range("E16").Value <> 0 Or _
           Not OpenSolver.IsZero(SensitivitySheet.Range("G13").Value - 21.6) Or _
           SensitivitySheet.Range("G14").Value <> 0 Or _
           SensitivitySheet.Range("G15").Value <> 0 Or _
           SensitivitySheet.Range("H13").Value <> 0 Or _
           Not OpenSolver.IsZero(SensitivitySheet.Range("H14").Value - 166.66666667) Or _
           SensitivitySheet.Range("H15").Value <> 100 Or _
           SensitivitySheet.Range("H16").Value <> 500 Then
            Test = Fail
        End If
    
        Application.DisplayAlerts = False
        SensitivitySheet.Delete
        Application.DisplayAlerts = True
    End If
    Exit Function
    
ErrorHandler:
    Test = Fail
End Function


