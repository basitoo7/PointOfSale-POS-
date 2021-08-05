VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTaxSchedule 
   Caption         =   "Tax Schedule"
   ClientHeight    =   2445
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
   Icon            =   "frmTaxSchedule.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
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
      Height          =   1860
      Left            =   15
      TabIndex        =   1
      Top             =   570
      Width           =   4920
      Begin VB.ComboBox txtType 
         Height          =   330
         ItemData        =   "frmTaxSchedule.frx":030A
         Left            =   1335
         List            =   "frmTaxSchedule.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1005
         Width           =   1410
      End
      Begin VB.TextBox txtfixamount 
         Height          =   315
         Left            =   3660
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1425
         Width           =   1140
      End
      Begin VB.TextBox txtperamount 
         Height          =   315
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1425
         Width           =   735
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFFF00&
         Height          =   330
         Left            =   1335
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "SKIPN"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   5
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
         Left            =   1875
         Picture         =   "frmTaxSchedule.frx":0328
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   210
         Left            =   2130
         TabIndex        =   11
         Top             =   1470
         Width           =   150
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fix Amount :"
         Height          =   210
         Left            =   2715
         TabIndex        =   9
         Top             =   1455
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Persentage :"
         Height          =   210
         Left            =   285
         TabIndex        =   8
         Top             =   1425
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Type :"
         Height          =   210
         Left            =   735
         TabIndex        =   7
         Top             =   1020
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   300
         TabIndex        =   4
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   735
         TabIndex        =   2
         Top             =   240
         Width           =   465
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
            Caption         =   "&Find"
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
               Picture         =   "frmTaxSchedule.frx":049A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTaxSchedule.frx":08EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTaxSchedule.frx":0D42
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTaxSchedule.frx":1196
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTaxSchedule.frx":15EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTaxSchedule.frx":1A3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTaxSchedule.frx":2192
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTaxSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_TaxSchedules As Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocation
    Set PO_DESC = txtdesc
    
    GoTop PR_TaxSchedules
    MyLookup.Caption = "Locations"
    MyLookup.FillGrid PR_TaxSchedules, "TaxCode", "Description", 5
    MyLookup.Show 1
    
    If Len(txtLocation) > 0 Then TxtLocation_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SRTAS00001")
  SetToolBar(2) = chkRights("SRTAS00002")
  SetToolBar(3) = chkRights("SRTAS00003")
  SetToolBar(4) = chkRights("SRTAS00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  
  Set PR_TaxSchedules = New Recordset
   
  PR_TaxSchedules.Open "Select * from IC_TaxSchedules where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
   
  PB_BlnkLoca = IIf(PR_TaxSchedules.EOF, True, False)

End Sub
Private Function maxtranscode() As String
Dim pr_dumy As New Recordset
pr_dumy.Open "select max(TaxCode) as transcode from IC_TaxSchedules where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
pr_dumy.Close
End Function

Private Sub Form_Unload(Cancel As Integer)
    PR_TaxSchedules.Close
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtType.SetFocus
End Sub
Private Sub txttype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtperamount.SetFocus
End Sub
Private Sub txtperamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtfixamount.SetFocus
End Sub



Private Sub TxtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) Then
         
         txtLocation.Text = IIf(IsNumeric(txtLocation.Text), DoPad(UCase(txtLocation.Text), txtLocation.MaxLength), UCase(txtLocation.Text))
         lb_found = MySeek(txtLocation.Text, "TaxCode", PR_TaxSchedules)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                   Call ClearVal
                   txtLocation.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                   Call ClearVal
                   txtLocation.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtLocation.Enabled = False
                      txtdesc.SetFocus
                   End If
                End If
            End Select
            
       End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
      txtLocation = maxtranscode
      txtLocation.Locked = True
    Else
      cmdLookup.Enabled = True
      txtLocation.Locked = False
    End If
    
    If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_TaxSchedules, Me, txtLocation, txtdesc, "X", "CompCount", 3, "TaxCode", "Description", 1, False, Toolbar1)
    End If
    If Mode = "A" Then txtdesc.SetFocus
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim ls_sql As String
PB_BlnkLoca = False
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              ls_sql = "INSERT into IC_TaxSchedules(compcode,TaxCode,Description,type,peramount,fixamount) VALUES ('" & Gs_compcode & "','" & txtLocation.Text & "','" & RepApp(txtdesc.Text) & "'," & txtType.ListIndex & "," & txtperamount & "," & txtfixamount & "  )"
              gc_dbcon.Execute ls_sql
              
           Case "E"
              ls_sql = "UPDATE IC_TaxSchedules SET Description= '" & RepApp(txtdesc.Text) & "', type = " & txtType.ListIndex & ", peramount = " & txtperamount & ", fixamount = " & txtfixamount & " WHERE  compcode = '" & Gs_compcode & "' and TaxCode= '" & txtLocation.Text & "'"
              gc_dbcon.Execute ls_sql
             
           Case "D"
              ls_sql = "DELETE FROM IC_TaxSchedules WHERE TaxCode = '" & txtLocation.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
           
     End Select

gc_dbcon.CommitTrans
PR_TaxSchedules.Requery
If Mode = "A" Then txtLocation = maxtranscode


Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Public Sub ClearVal()
     txtLocation = ""
     txtdesc = ""
End Sub

Private Sub SetVal()
     txtdesc = PR_TaxSchedules("Description")
     txtType.ListIndex = PR_TaxSchedules("type")
     txtperamount = PR_TaxSchedules("PerAmount")
     txtfixamount = PR_TaxSchedules("FixAmount")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtLocation.Text) = txtLocation.MaxLength And txtdesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtdesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub
