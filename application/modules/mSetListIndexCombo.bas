Attribute VB_Name = "mSetListIndexCombo"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186



Public Sub SetListIndex(vStrItem As String, vCombo As Control)
    ' Sets the ListIndex of a ComboBox or ListBox without firing it's click event.
    
    Dim i As Integer
    For i = 0 To vCombo.ListCount - 1
        If vCombo.List(i) = vStrItem Then
            Call SendMessage(vCombo.hwnd, CB_SETCURSEL, i, 0&)
            Exit Sub
        End If
    Next
    
    'If TypeOf vCombo Is ListBox Then
        'Call SendMessage(Ctrl.hwnd, LB_SETCURSEL, i, 0&)
    'ElseIf TypeOf vCombo Is ComboBox Then
        'Call SendMessage(vCombo.hwnd, CB_SETCURSEL, i, 0&)
    'End If
    
End Sub




