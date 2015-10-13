Attribute VB_Name = "FailWhitelist"
' A whitelist for tests that are expected to fail
Function TestShouldFail(SheetName As String, Solver As String) As Boolean
    Select Case SheetName & "_" & Solver
    
    ' ==================================
    ' COUENNE
    ' ==================================
    
    Case "InfConstConstraint_Couenne"       ' Reports optimal when problem is infeasible, bug with AMPL Couenne 0.5.1, NEOS versions work fine
        TestShouldFail = True
        
    Case "BinLB_Couenne"      ' Incorrect optimal solution, bound on binary var is removed. Bug with AMPL Couenne 0.5.1, NEOS versions work fine
        TestShouldFail = True
        
    Case "BinIntLB_Couenne"      ' Sometimes gets incorrect optimal solution, bound on binary var is removed. Bug with AMPL Couenne 0.5.1, NEOS versions work fine
        TestShouldFail = True
        
    Case "FormulaLB_Couenne"       ' Ignores lower bound, bug with AMPL Couenne 0.5.1, NEOS versions work fine
        TestShouldFail = True
        
    Case "SolverParameters_Couenne"       ' Error message on invalid option value is not redirected to log
        TestShouldFail = True
        
    Case "VarConstraintLB_Couenne"       ' Ignores lower bound, bug with AMPL Couenne 0.5.1, NEOS versions work fine
        TestShouldFail = True
        
    Case "Unbounded_Couenne"     ' Reports large optimal solution rather than unbounded
        TestShouldFail = True
        
    Case "NonLinMinMax_Couenne"       ' Couenne doesn't yet support MAX (0.5.1)
        TestShouldFail = True
        
    ' ==================================
    ' BONMIN
    ' ==================================
        
    ' ==================================
    ' NOMAD
    ' ==================================
        
    Case "Unbounded_NOMAD"  ' Reports large optimal solution rather than unbounded
        TestShouldFail = True
        
    ' ==================================
    ' NEOS COUENNE
    ' ==================================
    
    Case "Unbounded_NeosCou"     ' Reports large optimal solution rather than unbounded
        TestShouldFail = True

    Case "NonLinMinMax_NeosCou"       ' Couenne doesn't yet support MAX (0.4.7)
        TestShouldFail = True
        
    ' ==================================
    ' NEOS BONMIN
    ' ==================================
    
    ' ==================================


    Case Else
        TestShouldFail = False

    End Select
End Function
