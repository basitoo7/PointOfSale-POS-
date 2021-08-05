VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtehseel 
   Caption         =   "Tehseel Setup"
   ClientHeight    =   1995
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
   Icon            =   "frmtehseel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
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
      Height          =   1395
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   4935
      Begin VB.TextBox txttehseelname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2205
         MaxLength       =   25
         TabIndex        =   11
         Tag             =   "SKIP"
         Top             =   615
         Width           =   2595
      End
      Begin VB.TextBox txtcityname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2205
         MaxLength       =   25
         TabIndex        =   10
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   2595
      End
      Begin VB.CommandButton cmdlookup2 
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
         Left            =   1875
         Picture         =   "frmtehseel.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   615
         Width           =   315
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   975
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
         Left            =   1845
         Picture         =   "frmtehseel.frx":047C
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
         Width           =   510
         _ExtentX        =   900
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
      Begin MSMask.MaskEdBox txttehseel 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   615
         Width           =   510
         _ExtentX        =   900
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tehseel Code :"
         Height          =   210
         Left            =   195
         TabIndex        =   9
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "City Code :"
         Height          =   210
         Left            =   495
         TabIndex        =   6
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   375
         TabIndex        =   5
         Top             =   975
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
               Picture         =   "frmtehseel.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtehseel.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtehseel.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtehseel.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtehseel.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtehseel.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtehseel.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmtehseel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Cities As New Recordset
Dim PR_Tehseels As New Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcities
    Set PO_DESC = txtcityname
    
    GoTop PR_Cities
    MyLookup.Caption = "Cities"
    MyLookup.FillGrid PR_Cities, "CityCode", "CityName", txtcities.MaxLength
    MyLookup.Show 1
    
    If Len(txtcities) > 0 Then Txtcities_KeyDown vbKeyReturn, vbKeyShift

End Sub
Private Sub cmdlookup2_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttehseel
    Set PO_DESC = txttehseelname
    
    GoTop PR_Tehseels
    PR_Tehseels.Filter = "CityCode = '" & txtcities & "'"
    MyLookup.Caption = "Tehseels"
    MyLookup.FillGrid PR_Tehseels, "tehseelCode", "tehseelName", txttehseel.MaxLength
    MyLookup.Show 1
    PR_Tehseels.Filter = adFilterNone
    If Len(txttehseel) > 0 Then txttehseel_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Cities, Me, txtcities, txtdesc, "X", "CompCount", 3, "CityCode", "CityName", 1, False, Toolbar1)
End Sub

Private Sub Form_Load()
  SetToolBar(1) = chkRights("TEHSEEL001")
  SetToolBar(2) = chkRights("TEHSEEL002")
  SetToolBar(3) = chkRights("TEHSEEL003")
  SetToolBar(4) = chkRights("TEHSEEL004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
'  PR_GlDetail.Open "Select AccountNo,Acct_Desc FROM GL_detail Where Acct_Type ='B' and compcode = '" & Gs_compcode & "' Order By AccountNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_Cities.Open "Select * from Cities Order By CityCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Tehseels.Open "Select *,citycode+tehseelcode as Findfld from tehseels Order By CityCode,Tehseelcode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
 ' PR_VchType.Open "Select * from GlVchrType where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  
  PB_BlnkLoca = IIf(PR_Tehseels.EOF, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Cities.Close
    PR_Tehseels.Close
End Sub


Private Sub txtDesc_Change()
If txtdesc <> "" Then
  txttehseelname = txtdesc
End If
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageUp Then txtcities.SetFocus
End Sub

Private Sub Txtcities_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtcities.Text = DoPad(txtcities.Text, txtcities.MaxLength)
      If MySeek(txtcities.Text, "CityCode", PR_Cities) Then
           txtcityname = PR_Cities("Cityname")
           If Mode = "A" Then
               txttehseel.Enabled = False
               txttehseel = maxtranscode
               txtdesc.SetFocus
           Else
              txttehseel.Enabled = True
              txttehseel.SetFocus
               
           End If
        Else
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtcities.SetFocus
      End If
ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
End If
End Sub
Private Function maxtranscode() As String
Dim pr_dumy As New Recordset
pr_dumy.Open "select max(TehseelCode) as transcode from Tehseels where Citycode = '" & txtcities & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
pr_dumy.Close
End Function
Private Sub txttehseel_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txttehseel.Text = DoPad(txttehseel.Text, txttehseel.MaxLength)
      lb_found = MySeek(Trim(txtcities) + Trim(txttehseel.Text), "findfld", PR_Tehseels)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txttehseel.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txttehseel.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                     ' txtLocation.Enabled = False
                     txttehseelname = PR_Tehseels("tehseelname")
                      txtdesc.SetFocus
                   End If
                End If
            End Select
ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdlookup2.Enabled = False
      txttehseel.Enabled = False
    Else
      txttehseel.Enabled = True
      cmdlookup2.Enabled = True
    End If
    
  If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Cities, Me, txtcities, txtdesc, "X", "CompCount", 3, "CityCode", "CityName", 1, False, Toolbar1)
    End If
    If txtcities <> "" Then
     txttehseel = maxtranscode
     If txtdesc.Enabled Then txtdesc.SetFocus
    End If
End Sub

Public Sub SaveValues()
PB_BlnkLoca = False
'Dim ls_btype As String

'ls_btype = IIf(chklocal, "L", "F")
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT into tehseels(CityCode,tehseelcode,tehseelname) VALUES ('" & txtcities.Text & "','" & txttehseel.Text & "','" & txtdesc.Text & "')"
              
           Case "E"
              gc_dbcon.Execute "UPDATE tehseels SET tehseelName= '" & txtdesc.Text & "' WHERE  CityCode= '" & txtcities.Text & "' and tehseelcode = '" & txttehseel.Text & "'"
              
           Case "D"
              gc_dbcon.Execute "DELETE FROM tehseels  WHERE  CityCode = '" & txtcities.Text & "' And tehseelcode = '" & txttehseel.Text & "'"
     End Select
PR_Cities.Requery
PR_Tehseels.Requery

If Mode = "A" Then
 txttehseel.Enabled = False
 txttehseel = maxtranscode
 txtdesc.SetFocus
End If
End Sub

Private Sub SetVal()
     txtdesc = Trim(PR_Tehseels("tehseelName") & "")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtcities.Text) = txtcities.MaxLength And Len(txttehseel.Text) = txttehseel.MaxLength And txtdesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  PR_Cities.Requery
  PR_Tehseels.Requery
End Sub

