Attribute VB_Name = "Ctrl_PayRoll"
Public Sub Deplode(frm As Form)
    Dim X As Long
    Dim factor  As Double
    Dim Width As Integer
    Dim Height As Integer

    Height = frm.Height
    Width = frm.Width
    factor = Height / Width
'    frm.Width = 0
'    frm.Height = 0
'    frm.Show
    
    For X = Width To 0 Step -50
        frm.Width = X
        frm.Height = X * factor
        With frm
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 3
        End With
    Next X
    Unload frm
End Sub

Public Sub Explode(frm As Form)
    Dim X As Long
    Dim factor  As Double
    Dim Width As Integer
    Dim Height As Integer
    
    Height = frm.Height
    Width = frm.Width
    factor = Height / Width
    frm.Width = 0
    frm.Height = 0
    frm.Show
    
    For X = 0 To Width Step 50
        frm.Width = X
        frm.Height = X * factor
        With frm
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 3
        End With
    Next X
End Sub


Public Sub Populate_Text_Clear(frm As Form)
    For Each oCtrl In frm
        If TypeOf oCtrl Is TextBox Then oCtrl.Text = ""
        If TypeOf oCtrl Is MaskEdBox Then oCtrl.Text = "__/__/____"
        If TypeOf oCtrl Is ComboBox Then oCtrl.Clear
    Next
End Sub

Public Sub Populate_AutoID(Rst As ADODB.Recordset)
    With Rst
        If .RecordCount <= 0 Then
            Auto_ID = .RecordCount + 1
        ElseIf .RecordCount > 0 Then
            Auto_ID = .RecordCount + 1
        End If
    End With
End Sub

Public Sub Populate_CheckList(Lst As ListView, txt As TextBox)
    With Lst: Find_Flag = True
        For IntI = 1 To .ListItems.Count
            Set LItem = Lst.ListItems.Item(IntI)
                If txt.Text = LItem.SubItems(1) Then
                    MsgBox "You already have current Entery" & vbCrLf & _
                    "maust make it Unique", vbCritical, "List Entery Error!:"
                    Find_Flag = True: Exit Sub
                End If
        Next: Find_Flag = False
    End With
End Sub

Public Sub Populate_NumOnly(Asc As Integer)
Char = Chr(Asc)
    If Char Like "[0-9,.,__,.,/]" Or Asc = 8 Or Asc = 32 Or Asc = 13 Then
    Else: Asc = 0
    End If
End Sub

Public Sub Populate_CharOnly(Asc As Integer)
    Char = Chr(Asc)
    If Char Like "[a-z,.,A-Z,.,__,.,--,.,/]" Or Asc = 8 Or Asc = 32 Or Asc = 13 Then
    Else: Asc = 0
    End If
End Sub

Public Sub Populate_Alpha_Char(Asc As Integer, frm As Form, txt As TextBox)
    If ((Asc <> 13) And (Asc <> 8)) Then
        If Len(txt) = 0 Then Chrtxt = Asc: Asc = 0: _
           txt = UCase(Chr(Chrtxt)): SendKeys "{End}": Chrtxt = ""
        If Chrtxt = 32 Then txtVal = txt: Chrtxt = Asc: Asc = 0: _
           txt = txtVal & UCase(Chr(Chrtxt)): SendKeys "{End}": Chrtxt = ""
        If ((Asc = 32) And (Asc <> 8)) Then Chrtxt = Asc
    End If
End Sub

Public Sub Populate_Entery(frm As Form, Allow As Boolean)
    Dim oCtrl As Control
    For Each oCtrl In frm
        If TypeOf oCtrl Is TextBox Then oCtrl.Enabled = Allow
        If TypeOf oCtrl Is DTPicker Then oCtrl.Enabled = Allow
    Next
End Sub

Public Sub Populate_Init_Cmb(Rs As ADODB.Recordset, FldNo As Integer, Cmb As ComboBox)
    Cmb.Clear
    If Cmb.List(0) <> "Choose" Then Cmb.AddItem "Choose"
    With Rs
        If Rs.RecordCount > 0 Then
            .MoveFirst ': Cmb.Clear
            Do While Not .EOF
                If IsNull(.Fields(FldNo).Value) = False Then _
                    Cmb.AddItem .Fields(FldNo).Value
                .MoveNext
            Loop
        End If
    End With: Cmb.Text = "Choose"
End Sub

Public Sub msg_Consutruct()
    MsgBox "Under Construction: Please wait ...... ", vbInformation, "Under Construction"
End Sub
