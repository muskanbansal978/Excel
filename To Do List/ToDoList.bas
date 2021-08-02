Attribute VB_Name = "Module1"
Sub Complete()
Attribute Complete.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Complete Macro
'

'
    If ActiveCell.Value <> "" Then
        If ActiveCell.Offset(0, -1).Value <> 1 Then
            ActiveCell.Offset(0, -1).Value = 1
        Else
            ActiveCell.Offset(0, -1).Value = -1
        End If
    End If
End Sub
