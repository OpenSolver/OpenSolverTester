VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Warmstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    On Error GoTo ErrorHandler

    If Solver <> "Gurobi" Then
        Test = NA
        Exit Function
    End If

    OpenSolver.ResetModel Me
    OpenSolver.SetDecisionVariables Union(Me.Range("D1:UE1"), Me.Range("C23:C24")), sheet:=Me
    OpenSolver.SetObjectiveFunctionCell Me.Range("UF2"), sheet:=Me
    OpenSolver.SetObjectiveSense OpenSolver.ObjectiveSenseType.MaximiseObjective, sheet:=Me
    OpenSolver.AddConstraint Me.Range("UF4"), RelationLE, Me.Range("UH4"), sheet:=Me
    OpenSolver.AddConstraint Me.Range("D1:UE1"), RelationBIN, sheet:=Me
    
    OpenSolver.SetChosenSolver Solver, sheet:=Me
    OpenSolver.SetSolverParameters "Gurobi", Me.Range("C20:D20"), sheet:=Me
    OpenSolver.SetTolerance 0, sheet:=Me
    
    LoadOptimalSolution
    
    ' First try from the optimal solution
    RunTest ExpectOptimal:=True
    
    Dim VarCell As Range
    
    ' =========================================================
    ' D1 is a binary cell at value 1
    Set VarCell = Me.Range("D1")
    
    ' 1. Try setting it to slightly over 1
    VarCell.Value = 1 + OpenSolver.EPSILON / 2
    RunTest ExpectOptimal:=True
    
    ' 2. Try setting it to too far over 1
    VarCell.Value = 1 + OpenSolver.EPSILON * 2
    RunTest ExpectOptimal:=False
    
    ' 3. Try setting it to something non-integral
    VarCell.Value = 0.5
    RunTest ExpectOptimal:=False
    
    
    ' =========================================================
    ' G1 is a binary cell at value 0
    Set VarCell = Me.Range("G1")
    
    ' 1. Try setting it to slightly under 0
    VarCell.Value = 0 - OpenSolver.EPSILON / 2
    RunTest ExpectOptimal:=True
    
    ' 2. Try setting it to too far under 0
    VarCell.Value = 0 - OpenSolver.EPSILON * 2
    RunTest ExpectOptimal:=False
    
    ' 3. Try setting it to something non-integral
    VarCell.Value = 0.5
    RunTest ExpectOptimal:=False
    
    ' 4. Try setting it to 1, which violates the main constraint
    VarCell.Value = 1
    RunTest ExpectOptimal:=False
    
    ' ==========================================================
    ' Now play with the extra variable cell
    Set VarCell = Me.Range("C23")
    
    ' We start with assume non-neg vars on
    ' 1. Try slightly under zero
    VarCell.Value = 0 - OpenSolver.EPSILON / 2
    RunTest ExpectOptimal:=True
    ' 2. Try setting it to too far under 0
    VarCell.Value = 0 - OpenSolver.EPSILON * 2
    RunTest ExpectOptimal:=False
    
    ' Make it integer
    OpenSolver.AddConstraint VarCell, RelationINT, sheet:=Me
    ' 1. Set it to an integer value
    VarCell.Value = 5
    RunTest ExpectOptimal:=True
    ' 2. Set it to just under an integer value
    VarCell.Value = 5 - OpenSolver.EPSILON / 2
    RunTest ExpectOptimal:=True
    ' 3. Try setting it to too far under integer
    VarCell.Value = 5 - OpenSolver.EPSILON * 2
    RunTest ExpectOptimal:=False
    ' 4. Set it to just over an integer value
    VarCell.Value = 5 + OpenSolver.EPSILON / 2
    RunTest ExpectOptimal:=True
    ' 5. Try setting it to too far over integer
    VarCell.Value = 5 + OpenSolver.EPSILON * 2
    RunTest ExpectOptimal:=False
    ' Clean up
    OpenSolver.DeleteConstraint 3, sheet:=Me
    
    ' Add a lower bound
    OpenSolver.AddConstraint VarCell, RelationGE, RHSFormula:=2, sheet:=Me
    ' 1. Set it to the bound
    VarCell.Value = 2
    RunTest ExpectOptimal:=True
    ' 2. Set it to just under the bound
    VarCell.Value = 2 - OpenSolver.EPSILON / 2
    RunTest ExpectOptimal:=True
    ' 3. Try setting it to too far under the bound
    VarCell.Value = 2 - OpenSolver.EPSILON * 2
    RunTest ExpectOptimal:=False
    ' Clean up
    OpenSolver.DeleteConstraint 3, sheet:=Me
    
    ' Add a upper bound
    OpenSolver.AddConstraint VarCell, RelationLE, RHSFormula:=2, sheet:=Me
    ' 1. Set it to the bound
    VarCell.Value = 2
    RunTest ExpectOptimal:=True
    ' 2. Set it to just over the bound
    VarCell.Value = 2 + OpenSolver.EPSILON / 2
    RunTest ExpectOptimal:=True
    ' 3. Try setting it to too far over the bound
    VarCell.Value = 2 + OpenSolver.EPSILON * 2
    RunTest ExpectOptimal:=False
    ' Clean up
    OpenSolver.DeleteConstraint 3, sheet:=Me
    
    ' ========================================================
    ' Now we play with constraint violation
    ' This constraint involves two dummy variables so that it's not detected as bound
    Dim ConCell As Range, VarCell2 As Range
    Set ConCell = Me.Range("C25")
    Set VarCell2 = Me.Range("C24")
    
    ' Add a lower constraint
    OpenSolver.AddConstraint ConCell, RelationGE, RHSFormula:=2, sheet:=Me
    ' 1. Set it to the constraint
    VarCell.Value = 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=True
    ' 2. Set it to just under the constraint
    VarCell.Value = 2 - OpenSolver.EPSILON / 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=True
    ' 3. Try setting it to too far under the constraint
    VarCell.Value = 2 - OpenSolver.EPSILON * 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=False
    ' Clean up
    OpenSolver.DeleteConstraint 3, sheet:=Me
    
    ' Add a upper constraint
    OpenSolver.AddConstraint ConCell, RelationLE, RHSFormula:=2, sheet:=Me
    ' 1. Set it to the constraint
    VarCell.Value = 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=True
    ' 2. Set it to just over the constraint
    VarCell.Value = 2 + OpenSolver.EPSILON / 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=True
    ' 3. Try setting it to too far over the constraint
    VarCell.Value = 2 + OpenSolver.EPSILON * 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=False
    ' Clean up
    OpenSolver.DeleteConstraint 3, sheet:=Me
    
    ' Add an equality
    OpenSolver.AddConstraint ConCell, RelationEQ, RHSFormula:=2, sheet:=Me
    ' 1. Set it to the bound
    VarCell.Value = 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=True
    ' 2. Set it to just under the constraint
    VarCell.Value = 2 - OpenSolver.EPSILON / 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=True
    ' 3. Try setting it to too far under the constraint
    VarCell.Value = 2 - OpenSolver.EPSILON * 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=False
    ' 4. Set it to just over the constraint
    VarCell.Value = 2 + OpenSolver.EPSILON / 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=True
    ' 5. Try setting it to too far over the constraint
    VarCell.Value = 2 + OpenSolver.EPSILON * 2
    VarCell2.Value = 0
    RunTest ExpectOptimal:=False
    ' Clean up
    OpenSolver.DeleteConstraint 3, sheet:=Me
    
    Test = Pass
    Exit Function
    
ErrorHandler:
    Test = Fail
End Function

Sub LoadOptimalSolution()
' Load the optimal solution in, then reset the iteration limit.
    OpenSolver.SetMaxIterations 999999999, sheet:=Me
    If OpenSolver.RunOpenSolver(sheet:=Me, minimiseuserinteraction:=True) <> Optimal Or _
       Me.Range("A6").Value <> True Then
        Err.Raise 1
    End If
End Sub

Sub RunTest(ExpectOptimal As Boolean)
' Run the solve with the current values in the sheet and see if the warmstart was loaded
    OpenSolver.SetMaxIterations 10, sheet:=Me
    Dim SolveResult As OpenSolver.OpenSolverResult
    SolveResult = OpenSolver.RunOpenSolver(sheet:=Me, minimiseuserinteraction:=True)
    
    If ExpectOptimal Then
        If SolveResult <> Optimal Or Me.Range("A6").Value <> True Then
            Err.Raise 1
        End If
    Else
        If SolveResult = Optimal Then
            Err.Raise 1
        End If
    End If
    LoadOptimalSolution
End Sub
