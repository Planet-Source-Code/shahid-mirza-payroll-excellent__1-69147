Attribute VB_Name = "Ctrl_Variable"
Global DB_Conect As New ADODB.Connection 'For Open Database.
Global RstSQL As String 'For SQL Statement.

Global oCtrl As Control 'For Searching Control.
Global Opt_Priv, Opt_Stat 'For User Priviliges & Status Options.
Global Auto_ID As Integer 'For Generating Auto_ID Numbers.
Global Opt_Flag As String 'Flag For Acction Applied.
Global IntI As Integer 'For For Loop.
Global LItem As ListItem 'For List Items.
Global Find_Flag As Boolean 'For Record Find Flag.
Global LogIn_UID As String 'For Login User ID.
Global LogIn_Time As String 'For User Login Time.
Global UserName As String 'For Login User Name.
Global Catch_Field 'To Pick the Record from Database.
Global Req_Filed As String 'Required Field i.e Search.
Global Passport As String: Global CNIC As String
Global Chrtxt  'For Alpha Capital Character
Global Msg_Responce 'For Message Responce Yes/No/Cancel.

Public Sub Populate_LocalArea(Cmb As ComboBox)
    Cmb.Clear 'For Clearing the Combo Box.
    Cmb.AddItem "Choose":    Cmb.AddItem "Lahore"
    Cmb.AddItem "Karachi":   Cmb.AddItem "Islamabad"
    Cmb.Text = "Choose"
End Sub
