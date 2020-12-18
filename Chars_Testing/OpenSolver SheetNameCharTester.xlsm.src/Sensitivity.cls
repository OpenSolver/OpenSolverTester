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
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D4:F4"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("G6"), Me
    OpenSolver.AddConstraint Me.Range("G8"), RelationLE, Me.Range("I8"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("G9"), RelationGE, Me.Range("I9"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("G10:G11"), RelationLE, Me.Range("I10:I11"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("G20"), RelationEQ, Me.Range("G20"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    OpenSolver.SetObjectiveSense MaximiseObjective, Me
    OpenSolver.SetDualsOnSheet True, Me
    OpenSolver.SetUpdateSensitivity True, Me
    OpenSolver.SetDuals Me.Range("K1"), Me
    
    Test = RunTest(Me, Solver)
    If SensitivityAnalysisAvailable(CreateSolver(Solver)) Then
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
        
        If Me.Range("K1").Value <> "Cons" Or _
           Me.Range("L1").Value <> "SP" Or _
           Me.Range("M1").Value <> "Inc" Or _
           Me.Range("N1").Value <> "Dec" Then
            Test = Fail
        End If
        
        If Me.Range("K2").Value <> "G8<=I8" Or _
           Me.Range("K3").Value <> "G9>=I9" Or _
           Me.Range("K4").Value <> "G10<=I10" Or _
           Me.Range("K5").Value <> "G11<=I11" Or _
           Me.Range("K6").Value <> "G20=G20" Then
            Test = Fail
        End If

        If Me.Range("L2").Value <> 0 Or _
           Me.Range("L3").Value <> -1 Or _
           Me.Range("L4").Value <> 0 Or _
           Me.Range("L5").Value <> 2 Or _
           Me.Range("L6").Value <> 0 Or _
           Me.Range("M3").Value <> 3 Or _
           Me.Range("M5").Value <> 2 Or _
           Me.Range("M6").Value <> 0 Or _
           Me.Range("N2").Value <> 5 Or _
           Me.Range("N3").Value <> 2 Or _
           Me.Range("N4").Value <> 2 Or _
           Me.Range("N5").Value <> 3 Or _
           Me.Range("N6").Value <> 0 Then
            Test = Fail
        End If
        
        If Me.Range("K8").Value <> "Vars" Or _
           Me.Range("L8").Value <> "RC" Or _
           Me.Range("M8").Value <> "Inc" Or _
           Me.Range("N8").Value <> "Dec" Then
            Test = Fail
        End If
        
        If Me.Range("K9").Value <> "D4" Or _
           Me.Range("K10").Value <> "E4" Or _
           Me.Range("K11").Value <> "F4" Then
            Test = Fail
        End If
        
        If Me.Range("L9").Value <> -1 Or _
           Me.Range("L10").Value <> 0 Or _
           Me.Range("L11").Value <> 0 Or _
           Me.Range("M9").Value <> 1 Or _
           Not OpenSolver.IsZero(Me.Range("M10").Value - 1) Or _
           Not OpenSolver.IsZero(Me.Range("N11").Value - 1) Then
            Test = Fail
        End If
    End If
    Exit Function
    
ErrorHandler:
    Test = Fail
End Function


