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
    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Me.Range("D13:F13"), Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("H15"), Me
    OpenSolver.AddConstraint Me.Range("H17:H18"), RelationLE, Me.Range("J17:J18"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H19"), RelationGE, Me.Range("J19"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("H20"), RelationEQ, Me.Range("J20"), sheet:=Me
    OpenSolver.SetChosenSolver Solver, Me
    OpenSolver.SetNonNegativity False, Me
    
    OpenSolver.DeleteSolverParameters Solver, Me
    
    Dim Offset As Long, ParameterName As String
    Select Case Solver
    Case "CBC", "NeosCBC"
        OpenSolver.SetMaxIterations 1.5, Me
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.LimitedSubOptimal Then GoTo Failed
        
        ' Reset MaxIterations
        OpenSolver.SetMaxIterations 9999999, Me
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.Unbounded Then GoTo Failed

        ' Test custom parameters for CBC
        Offset = IIf(Solver = "CBC", 0, 6)
        Me.Cells(4, 3 + Offset).Value = Solver & " Options"
        OpenSolver.SetSolverParameters Solver, Me.Range(IIf(Solver = "CBC", "C5:D7", "I5:J7")), Me
        
        ' Bad parameter name
        ParameterName = "duBound"
        Me.Cells(5, 3 + Offset).Value = ParameterName
        Me.Cells(5, 4 + Offset).Value = 0.21
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        Me.Cells(5, 3 + Offset).Value = "dualBound"  ' Restore name
        
        ' Bad parameter value
        ParameterName = "forceSolution"
        Me.Cells(7, 3 + Offset).Value = ParameterName
        Me.Cells(7, 4 + Offset).Value = 0.5
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Valid parameters
        Me.Cells(5, 3 + Offset).Value = "dualBound"
        Me.Cells(6, 3 + Offset).Value = "dualTolerance"
        Me.Cells(7, 3 + Offset).Value = "forceSolution"
        Me.Cells(5, 4 + Offset).Value = 0.21
        Me.Cells(6, 4 + Offset).Value = 0.7
        Me.Cells(7, 4 + Offset).Value = 1
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.Optimal Or _
           Me.Range("A6").Value <> True Then GoTo Failed
           
    Case "Gurobi"
        OpenSolver.SetMaxIterations 1.5, Me
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.LimitedSubOptimal Then GoTo Failed
        
        ' Test custom parameters for Gurobi
        Me.Range("F4").Value = "Gurobi Options"
        OpenSolver.SetSolverParameters "Gurobi", Me.Range("F5:G5"), Me
        
        ' Bad parameter name
        ParameterName = "IteratLimit"
        Me.Range("F5").Value = ParameterName
        Me.Range("G5").Value = 10
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Bad parameter value
        ParameterName = "IterationLimit"
        Me.Range("F5").Value = ParameterName
        Me.Range("G5").Value = -1
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Valid parameters
        Me.Range("F5").Value = "IterationLimit"
        Me.Range("G5").Value = 9999999
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.Unbounded Then GoTo Failed
        
    Case "Bonmin", "NeosBon"
        OpenSolver.GetDecisionVariables(Me).Value = 0
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.Unbounded Then GoTo Failed
        
        ' Test custom parameters for Bonmin
        Offset = IIf(Solver = "Bonmin", 0, 9)
        Me.Cells(4, 12 + Offset).Value = Solver & " Options"
        OpenSolver.SetSolverParameters Solver, Me.Range(IIf(Solver = "Bonmin", "L5:M5", "U5:V5")), Me
        
        ' Bad parameter name
        ParameterName = "m_iter"
        Me.Cells(5, 12 + Offset).Value = ParameterName
        Me.Cells(5, 13 + Offset).Value = 1
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Bad parameter value
        ParameterName = "max_iter"
        Me.Cells(5, 12 + Offset).Value = ParameterName
        Me.Cells(5, 13 + Offset).Value = -1
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Valid parameters
        Me.Cells(5, 12 + Offset).Value = "max_iter"
        Me.Cells(5, 13 + Offset).Value = 1
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) = OpenSolverResult.Unbounded Then GoTo Failed
        
    Case "Couenne", "NeosCou"
        OpenSolver.RunOpenSolver False, True, sheet:=Me
        ' Couenne can't detect unbounded here yet and returns optimal with a large objective
        If OpenSolver.GetObjectiveFunctionCell(Me).Value > -10000000000# Then GoTo Failed
        
        ' Test custom parameters for Couenne
        Offset = IIf(Solver = "Couenne", 0, 9)
        Me.Cells(4, 15 + Offset).Value = Solver & " Options"
        OpenSolver.SetSolverParameters Solver, Me.Range(IIf(Solver = "Couenne", "O5:P5", "X5:Y5")), Me
        
        ' Bad parameter name
        ParameterName = "t_limit"
        Me.Cells(5, 15 + Offset).Value = ParameterName
        Me.Cells(5, 16 + Offset).Value = 0.00001
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Bad parameter value
        ParameterName = "time_limit"
        Me.Cells(5, 15 + Offset).Value = ParameterName
        Me.Cells(5, 16 + Offset).Value = -1
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Valid parameters
        Me.Cells(5, 15 + Offset).Value = "time_limit"
        Me.Cells(5, 16 + Offset).Value = 0.00001
        Dim TempSolveResult As OpenSolverResult
        TempSolveResult = OpenSolver.RunOpenSolver(False, True, sheet:=Me)
        If TempSolveResult <> OpenSolverResult.Unbounded And TempSolveResult <> OpenSolverResult.Infeasible Then GoTo Failed
        
    Case "NOMAD"
        OpenSolver.GetDecisionVariables(Me).Value = 0
        Me.Range("E13").Value = 180 / 34
        OpenSolver.SetMaxIterations 20.5, Me
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.LimitedSubOptimal Then GoTo Failed
        
        ' Test custom parameters for NOMAD
        Me.Range("R4").Value = Solver & " Options"
        OpenSolver.SetSolverParameters Solver, Me.Range("R5:S5"), Me
        
        ' Bad parameter name
        ParameterName = "stif_feasible"
        Me.Range("R5").Value = ParameterName
        Me.Range("S5").Value = 1
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Bad parameter value
        ParameterName = "stop_if_feasible"
        Me.Range("R5").Value = ParameterName
        Me.Range("S5").Value = 0.5
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.ErrorOccurred Then GoTo Failed
        If Not InStr(OpenSolver.ErrMsg, ParameterName) > 0 Then GoTo Failed
        
        ' Valid parameters
        Me.Range("R5").Value = "stop_if_feasible"
        Me.Range("S5").Value = 1
        If OpenSolver.RunOpenSolver(False, True, sheet:=Me) <> OpenSolverResult.Optimal Then GoTo Failed
    Case Else
        GoTo Failed
    End Select
    
    Test = Pass
    Exit Function
    
Failed:
    Test = Fail
End Function

