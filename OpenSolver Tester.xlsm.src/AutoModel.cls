VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function Test(Solver As String) As TestResult
    OpenSolver.ResetModel Me
    OpenSolver.RunAutoModel Me, True
    
    Test = Pass
    
    ' Check objective
    If OpenSolver.GetObjectiveFunctionCell(Me).Address <> Me.Range("C9").Address Then
        Test = Fail
        Exit Function
    End If
    
    ' Check decision vars
    Dim Intersection As Range
    Set Intersection = Intersect(OpenSolver.GetDecisionVariables(Me), Me.Range("C3:C5,B28:D28"))
    If Intersection Is Nothing Then
        Test = Fail
        Exit Function
    ElseIf Intersection.Count <> OpenSolver.GetDecisionVariables(Me).Count Then
        Test = Fail
        Exit Function
    End If
    
    ' Check constraints present
    If Not FindConstraint(Me.Range("B14"), RelationGE, Me.Range("D14")) Or _
       Not FindConstraint(Me.Range("B17"), RelationGE, Me.Range("D17")) Or _
       Not FindConstraint(Me.Range("B20"), RelationLE, Me.Range("D20")) Or _
       Not FindConstraint(Me.Range("D33"), RelationGE, Me.Range("D35")) Or _
       Not FindConstraint(Me.Range("B28:D28"), RelationINT) Then
        Test = Fail
    End If

End Function

Function FindConstraint(LHS As Range, rel As RelationConsts, Optional RHS As Range) As Boolean
    Dim i As Long
    For i = 1 To OpenSolver.GetNumConstraints(Me)
        If LHS.Address = OpenSolver.GetConstraintLhs(i, Me).Address And _
           rel = OpenSolver.GetConstraintRel(i) Then
            If Not RHS Is Nothing Then
                If RHS.Address = OpenSolver.GetConstraintRhs(i, "", 0, False, Me).Address Then
                    FindConstraint = True
                    Exit Function
                End If
            Else
                FindConstraint = True
                Exit Function
            End If
        End If
    Next i
    FindConstraint = False
End Function
