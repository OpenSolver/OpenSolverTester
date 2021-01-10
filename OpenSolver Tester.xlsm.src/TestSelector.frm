VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestSelector 
   Caption         =   "Select tests to run"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   -2220
   ClientWidth     =   7068
   OleObjectBlob   =   "TestSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub RefreshTestListBox()
    With Me.lstTests
        .Clear
    
        ' Loop through worksheets looking for sheets with the testing pane present.
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ' Exit if not a testing sheet
            If ws.Cells(2, 1).Value = "Normal" Or ws.Cells(2, 1).Value = "Custom" Then
                .AddItem ws.Name
            End If
        Next ws
    End With
    ' Decheck select all box
    Me.chkAllTests.Value = False
End Sub

Private Sub chkAllTests_Change()
    With Me.lstTests
        If Me.chkAllTests.Value = True Then
            Dim i As Integer
            For i = 0 To .ListCount - 1
                .Selected(i) = True
            Next i
            .Enabled = False
        Else
            .Enabled = True
        End If
    End With
End Sub


Private Sub cmdClear_Click()
    Me.chkAllTests.Value = False
    With Me.lstTests
        Dim i As Long
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
    End With
End Sub

Private Sub cmdRun_Click()
    Me.Hide
    SOLVE_LOCAL = chkLocalNeos.Value
    RunAllTests chkClearResults.Value
    SOLVE_LOCAL = False
End Sub

Private Sub UserForm_Initialize()
    RefreshTestListBox
    RefreshSolvers
End Sub

Sub RefreshSolvers()
    Dim SolverShortName As Variant, Solver As Object
    
    Me.lstLinearSolvers.Clear
    Me.lstNonLinearSolvers.Clear
    For Each SolverShortName In OpenSolver.GetAvailableSolvers()
        Set Solver = OpenSolver.CreateSolver(CStr(SolverShortName))
        If SolverLinearity(Solver) = OpenSolver.OpenSolver_SolverType.Linear Then
            Me.lstLinearSolvers.AddItem Solver.ShortName
        Else
            Me.lstNonLinearSolvers.AddItem Solver.ShortName
        End If
    Next SolverShortName
End Sub

