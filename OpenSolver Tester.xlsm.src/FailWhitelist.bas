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
        
    Case "VarConstraintLB_Couenne"       ' Ignores lower bound, bug with AMPL Couenne 0.5.1, NEOS versions work fine
        TestShouldFail = True
        
    Case "Unbounded_Couenne"     ' Reports large optimal solution rather than unbounded
        TestShouldFail = True
        
    Case "NonLinMinMax_Couenne"       ' Couenne doesn't yet support MAX (0.5.1)
        TestShouldFail = True
        
    ' ==================================
    ' BONMIN
    ' ==================================
        
    Case "NonLin6_Bonmin"        ' Bonmin can't solve this problem from the given starting points
        TestShouldFail = True
        
    ' ==================================
    ' NOMAD
    ' ==================================
        
    Case "Unbounded_NOMAD"  ' Reports large optimal solution rather than unbounded
        TestShouldFail = True
        
    Case "NonLin6_NOMAD"  ' Sometimes fails based on starting solution
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
    
    Case "NonLin6_NeosBon"        ' Bonmin can't solve this problem from the given starting points
        TestShouldFail = True


    Case Else
        TestShouldFail = False

    End Select
End Function
