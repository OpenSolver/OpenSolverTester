Attribute VB_Name = "NOMADCallbackFunc"
Sub NOMADCallback_updateObjective()
    ' Updates C4 to be C6^2
    NOMADCallback.Cells(4, 4) = NOMADCallback.Cells(6, 4) ^ 2
End Sub
