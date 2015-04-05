VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sensitivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    Me.Activate
    OpenSolver.ResetModel
    OpenSolver.SetDecisionVariables Me.Range("D4:F4")
    OpenSolver.SetObjectiveFunctionCell Me.Range("G6")
    OpenSolver.AddConstraint Me.Range("G8"), RelationLE, Me.Range("I8")
    OpenSolver.AddConstraint Me.Range("G9"), RelationGE, Me.Range("I9")
    OpenSolver.AddConstraint Me.Range("G10:G11"), RelationLE, Me.Range("I10:I11")
    OpenSolver.SetChosenSolver Solver
    OpenSolver.SetObjectiveSense MaximiseObjective
    OpenSolver.SetDualsNewSheet True
    OpenSolver.SetUpdateSensitivity True
    
    Test = RunTest(Me, Solver)
    If SolverHasSensitivityAnalysis(Solver) Then
        Dim SensitivitySheet As Worksheet
        On Error GoTo ErrorHandler
        Set SensitivitySheet = Sheets(Me.Name & " Sensitivity")
        On Error GoTo 0
        
        If SensitivitySheet.Range("E7").Value <> -1 Or _
           SensitivitySheet.Range("E8").Value <> 0 Or _
           SensitivitySheet.Range("E9").Value <> 0 Or _
           SensitivitySheet.Range("G7").Value <> 1 Or _
           Not OpenSolver.IsZero(SensitivitySheet.Range("G8").Value - 1) Or _
           Not OpenSolver.IsZero(SensitivitySheet.Range("H9").Value - 1) Or _
           SensitivitySheet.Range("E13").Value <> 0 Or _
           SensitivitySheet.Range("E14").Value <> -1 Or _
           SensitivitySheet.Range("E15").Value <> 0 Or _
           SensitivitySheet.Range("E16").Value <> 2 Or _
           SensitivitySheet.Range("G14").Value <> 3 Or _
           SensitivitySheet.Range("G16").Value <> 2 Or _
           SensitivitySheet.Range("H13").Value <> 5 Or _
           SensitivitySheet.Range("H14").Value <> 2 Or _
           SensitivitySheet.Range("H15").Value <> 2 Or _
           SensitivitySheet.Range("H16").Value <> 3 Then
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


