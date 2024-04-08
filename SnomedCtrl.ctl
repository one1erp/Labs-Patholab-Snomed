VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SnomedCtrl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4080
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin MSComctlLib.ListView LstSnomed 
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdCode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   14111
      EndProperty
   End
End
Attribute VB_Name = "SnomedCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Connection As ADODB.Connection
Private rst As ADODB.Recordset
Private StatusWR As Integer
Private Const ReadOnly = 1
Private Const ReadWrite = 2
Private Const InsertButton = 1
Private Const DeleteButton = 2
Private Const CloseButton = 3
Public Event CloseClick()
Private Const NewSnomed = "New Snomed"
Dim WithEvents ClueBrw As ClueBrowser
Attribute ClueBrw.VB_VarHelpID = -1
Private SnomedHwnd As Long


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Sub Initialize(SQLStr As String)
    FillList (SQLStr)
    If LstSnomed.ListItems.Count > 0 Then
        LstSnomed.ListItems.Item(1).Selected = False
    End If
End Sub
Public Sub Terminate()
    LstSnomed.ListItems.Clear
    Set ClueBrw = Nothing
End Sub
Private Sub FillList(SQLStr As String)
    Dim Snomeds As String
    Dim Snomed As String
    Dim li As ListItem
    If SQLStr = "" Then
        Exit Sub
    End If
    LstSnomed.ListItems.Clear
    Set rst = Connection.Execute(SQLStr)
    If rst.EOF Then
        Snomeds = ""
    Else
        Snomeds = nte(rst(0))
    End If
    Do While Snomeds <> ""
        Snomed = getNextSnomed(Snomeds)
        Set li = LstSnomed.ListItems.Add(, , Snomed)
        li.SubItems(1) = ClueInit.ClueRef.ConceptExpand(Snomed)
    Loop
    If StatusReadWrite = CReadWrite Then
        LstSnomed.ListItems.Add , NewSnomed, NewSnomed
    End If
End Sub

Private Sub CmdClose_Click()
    Terminate
    RaiseEvent CloseClick
End Sub

Private Sub cmdDelete_Click()
    Delete
End Sub

Private Sub cmdInsert_Click()
    Insert
End Sub

Private Sub LstSnomed_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim SnomedDesc As String
    SnomedDesc = ClueInit.ClueRef.ConceptExpand(Trim(NewString))
    If SnomedDesc = "" Then
        MsgBox "Snomed does not exist"
        Cancel = True
    Else
        If LstSnomed.ListItems(LstSnomed.SelectedItem.Index).Selected Then
            LstSnomed.SelectedItem.SubItems(1) = SnomedDesc
            If LstSnomed.SelectedItem.key = NewSnomed Then
                LstSnomed.SelectedItem.key = ""
                LstSnomed.ListItems.Add , NewSnomed, NewSnomed
            End If
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    LstSnomed.Left = 0
    LstSnomed.Top = 0
    UserControl.Width = LstSnomed.Width
    UserControl.Height = LstSnomed.Height + cmdClose.Height + 100
End Sub

Public Function getFirstSnomed(SQLStr As String) As String
    Dim Snomeds As String
    If SQLStr = "" Then
        Exit Function
    End If
    Set rst = Connection.Execute(SQLStr)
    If rst.EOF Then
        Snomeds = ""
    Else
        Snomeds = nte(rst(0))
    End If
    If Not rst.EOF Then
        getFirstSnomed = getNextSnomed(Snomeds)
    Else
        getFirstSnomed = ""
    End If
End Function

Public Sub CloseSnomed()
    ClueInit.Quit
End Sub

Private Sub UserControl_Resize()
    LstSnomed.Left = 0
    LstSnomed.Top = 0
    LstSnomed.Width = UserControl.Width
    If UserControl.Height - cmdClose.Height - 100 > 0 Then
        LstSnomed.Height = UserControl.Height - cmdClose.Height - 100
    End If
    cmdClose.Top = LstSnomed.Top + LstSnomed.Height + 50
    cmdDelete.Top = LstSnomed.Top + LstSnomed.Height + 50
    cmdInsert.Top = LstSnomed.Top + LstSnomed.Height + 50
    UserControl.Height = cmdClose.Top + cmdClose.Height + 50
End Sub

Private Function getNextSnomed(ByRef Str As String) As String
    If InStr(1, Str, ",") = 0 Then
        getNextSnomed = Str
        Str = ""
        Exit Function
    End If
    getNextSnomed = Mid(Str, 1, InStr(1, Str, ",") - 1)
    Str = Mid(Str, InStr(1, Str, ",") + 1)
End Function

Public Property Get StatusReadWrite() As Integer
    StatusReadWrite = StatusWR
End Property

Public Property Let StatusReadWrite(ByVal vNewValue As Integer)
    StatusWR = vNewValue
    If StatusReadWrite = ReadOnly Then
       cmdClose.Left = 0
       cmdDelete.Visible = False
       cmdInsert.Visible = False
       LstSnomed.LabelEdit = lvwManual
    Else
       cmdClose.Left = cmdDelete.Left + cmdDelete.Width + 50
       cmdDelete.Visible = True
       cmdInsert.Visible = True
       LstSnomed.LabelEdit = lvwAutomatic
    End If
End Property

Public Property Get CReadOnly() As Integer
    CReadOnly = ReadOnly
End Property
Public Property Get CReadWrite() As Integer
    CReadWrite = ReadWrite
End Property

Public Sub ShowCloseBtn(Show As Boolean)
    cmdClose.Visible = Show
End Sub

Private Sub Insert()
'MsgBox 1
    If ClueBrw Is Nothing Then Set ClueBrw = ClueInit.ClueBrw
'MsgBox 2
    ClueInit.UserControlId = UserControl.hwnd
'MsgBox 3
    ClueBrw.Show
'MsgBox 4
    ClueBrw.WindowState = vbNormal
'MsgBox 5
    ClueBrw.SearchType = smNormal
'MsgBox 6
    ClueBrw.key = Note
'MsgBox 7
'    SnomedHwnd = FindWindow("ThunderRT6FormDC", vbNullString)
'MsgBox 8
'    ShowWindow SnomedHwnd, SW_SHOWMAXIMIZED
'MsgBox 9
End Sub
Private Sub Delete()
    If LstSnomed.ListItems(LstSnomed.SelectedItem.Index).Selected And _
        LstSnomed.SelectedItem.key <> NewSnomed Then
        LstSnomed.ListItems.Remove LstSnomed.SelectedItem.Index
    End If
    If LstSnomed.ListItems.Count > 0 Then
        LstSnomed.ListItems.Item(1).Selected = False
    End If
End Sub

Private Sub ClueBrw_ButtonClick(ButtonTag As String)
    Dim li As ListItem
    If ClueInit.UserControlId <> UserControl.hwnd Then Exit Sub
    Select Case ButtonTag
        Case "Ok"
            With ClueInit.ClueRef.Concept(ClueBrw.ConceptId)
                 Set li = LstSnomed.ListItems.Item(NewSnomed)
                 li.Text = .SnomedId
                 li.key = ""
                 li.SubItems(1) = ClueInit.ClueRef.ConceptExpand(.SnomedId)
            End With
            If LstSnomed.ListItems.Count > 0 Then
                LstSnomed.ListItems.Item(1).Selected = False
            End If
            LstSnomed.ListItems.Add , NewSnomed, NewSnomed
         Case "Cancel"
            ShowWindow SnomedHwnd, SW_SHOWMINIMIZED
    End Select
    ClueInit.UserControlId = 0
    ClueBrw.Hide
End Sub

Public Function getSnomeds() As String
    Dim i As Integer
    Dim Snomeds As String
    For i = 1 To LstSnomed.ListItems.Count - 1
        Snomeds = Snomeds + LstSnomed.ListItems.Item(i).Text
        If i < LstSnomed.ListItems.Count - 1 Then
            Snomeds = Snomeds & ","
        End If
    Next i
    getSnomeds = Snomeds
End Function

Private Function nte(e As Variant) As Variant
    nte = IIf(IsNull(e), "", e)
End Function

Public Property Get getParser() As Parser
    Set getParser = New Parser
End Property

