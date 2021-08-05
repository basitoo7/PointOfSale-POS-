VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmexptype 
   Caption         =   "Expense Types"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExptype.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   4935
      Begin VB.TextBox Textx 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   4620
         MaxLength       =   35
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton cmdLookup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1815
         Picture         =   "frmExptype.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtcities 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   375
         TabIndex        =   6
         Top             =   600
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            Description     =   "Add"
            Object.ToolTipText     =   "Add new record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            Description     =   "Edit"
            Object.ToolTipText     =   "Edit an existing record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            Description     =   "Remove "
            Object.ToolTipText     =   "Remove an existing record."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            Description     =   "Save a new Record"
            Object.ToolTipText     =   "Save on disk"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Listing"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Description     =   "Find a Record."
            Object.ToolTipText     =   "Find a record."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel"
            Description     =   "Cancel Operation"
            Object.ToolTipText     =   "Cancel operation mode"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   14
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4920
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExptype.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExptype.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExptype.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExptype.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExptype.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExptype.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExptype.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmexptype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Cities As New Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcities
    Set PO_DESC = txtDesc
    
    GoTop PR_Cities
    MyLookup.Caption = "Expense Types"
    MyLookup.FillGrid PR_Cities, "ExpCode", "ExpName", txtcities.MaxLength
    MyLookup.Show 1
    
    If Len(txtcities) > 0 Then Txtcities_KeyDown vbKeyReturn, vbKeyShift

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Cities, Me, txtcities, txtDesc, "X", "CompCount", 3, "CityCode", "CityName", 1, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("CITY000001")
  SetToolBar(2) = chkRights("CITY000002")
  SetToolBar(3) = chkRights("CITY000003")
  SetToolBar(4) = chkRights("CITY000004")
  
  Toolbar1.Buttons(1).Enabled = True
  Toolbar1.Buttons(2).Enabled = True
  Toolbar1.Buttons(3).Enabled = True
  Toolbar1.Buttons(5).Enabled = True
  
'  PR_GlDetail.Open "Select AccountNo,Acct_Desc FROM GL_detail Where Acct_Type ='B' and compcode = '" & Gs_compcode & "' Order By AccountNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_Cities.Open "Select * from Ic_Expense Order By ExpCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
 ' PR_VchType.Open "Select * from GlVchrType where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  
  PB_BlnkLoca = IIf(PR_Cities.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Cities.Close
End Sub


Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageUp Then txtcities.SetFocus
End Sub

Private Sub Txtcities_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtcities.Text = DoPad(txtcities.Text, txtcities.MaxLength)
      lb_found = MySeek(txtcities.Text, "ExpCode", PR_Cities)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtcities.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtcities.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                     ' txtLocation.Enabled = False
                      txtDesc.SetFocus
                   End If
                End If
            End Select
ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
    Else
      cmdLookup.Enabled = True
    End If
    
  If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Cities, Me, txtcities, txtDesc, "X", "CompCount", 3, "CityCode", "CityName", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
PB_BlnkLoca = False
'Dim ls_btype As String

'ls_btype = IIf(chklocal, "L", "F")
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT into IC_Expense(ExpCode,ExpName) VALUES ('" & txtcities.Text & "','" & txtDesc.Text & "')"
           Case "E"
              gc_dbcon.Execute "UPDATE IC_Expense SET ExpName= '" & txtDesc.Text & "' WHERE  ExpCode= '" & txtcities.Text & "'"
           Case "D"
              gc_dbcon.Execute "DELETE FROM IC_Expense WHERE ExpCode = '" & txtcities.Text & "'"
     End Select
PR_Cities.Requery
End Sub

Private Sub SetVal()
     txtDesc = Trim(PR_Cities("ExpName") & "")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtcities.Text) = txtcities.MaxLength And txtDesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  PR_Cities.Requery
End Sub

