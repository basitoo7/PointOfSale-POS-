VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBranches 
   Caption         =   "Company Branches"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBranches.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5475
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
      Width           =   5415
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
         Picture         =   "frmBranches.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "frmBranches.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         MaxLength       =   64
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   3075
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   3975
      End
      Begin MSMask.MaskEdBox txtProcess 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   600
         Width           =   525
         _ExtentX        =   926
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
      Begin MSMask.MaskEdBox txtcompcode 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
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
         AutoSize        =   -1  'True
         Caption         =   "Company Code :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Branch Code :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   9
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   7
         Top             =   960
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   1005
      ButtonWidth     =   1323
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
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
            Caption         =   "Re&fresh"
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
         Left            =   4080
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
               Picture         =   "frmBranches.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBranches.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBranches.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBranches.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBranches.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBranches.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBranches.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmBranches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object

Public PR_Company As Recordset
Dim PR_SyProc As Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtprocess
    Set PO_DESC = txtdesc
    GoTop PR_SyProc
    PR_SyProc.Filter = "Compcode = '" & txtCompCode & "'"
    
    MyLookup.Caption = "Processes"
    MyLookup.FillGrid PR_SyProc, "BranchCode", "BranchDesc", 3
    MyLookup.Show 1
    
    If Len(txtprocess) > 0 Then txtprocess_KeyDown vbKeyReturn, vbKeyShift
    PR_SyProc.Filter = adFilterNone
    
End Sub

Private Sub Command1_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCompCode
    Set PO_DESC = Text1
    GoTop PR_SyProc
    MyLookup.Caption = "Companies"
    MyLookup.FillGrid PR_Company, "CompCode", "CompName", 3
    MyLookup.Show 1
    
    If Len(txtCompCode) > 0 Then txtcompcode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then
    Mode = DentMode(Mode, 4, PR_SyProc, Me, txtCompCode, txtprocess, PR_Company, "BranchCnt", 3, "BranchCode", "BranchDesc", 1, False, Toolbar1)
End If
End Sub

Private Sub Form_Load()
  
' Setting up Preveliges
  
  SetToolBar(1) = chkRights("SMBRANCH01")
  SetToolBar(2) = chkRights("SMBRANCH02")
  SetToolBar(3) = chkRights("SMBRANCH03")
  SetToolBar(4) = chkRights("SMBRANCH04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Set PR_SyProc = New Recordset
  Set PR_Company = New Recordset
  
  PR_Company.Open "Select Compcode,Compname,BranchCnt from SysComp", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_SyProc.Open "Select *,Compcode+BranchCode As Findfld from SysBranch", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    
  PB_BlnkGls0 = IIf(PR_SyProc.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Company.Close
    PR_SyProc.Close
End Sub

Private Sub txtcompcode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
lb_found = False
If KeyCode = vbKeyReturn Then
    If txtCompCode <> "" Then
        txtCompCode = DoPad(txtCompCode, txtCompCode.MaxLength)
        lb_found = MySeek(txtCompCode.Text, "Compcode", PR_Company)
        
        If lb_found Then
           Text1 = PR_Company("CompName")
           txtprocess.SetFocus
        Else
           Call SetErr(Gs_RecNFMsg, vbCritical)
           txtCompCode.SetFocus
        End If
    End If
End If
End Sub

Private Sub txtprocess_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtprocess.Text) > 0 Then
          
         txtprocess.Text = DoPad(txtprocess.Text, txtprocess.MaxLength)
         lb_found = MySeek(txtCompCode + txtprocess.Text, "FindFld", PR_SyProc)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtprocess.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtprocess.SetFocus
                Else
                   Call SetVal
                   If txtdesc.Enabled Then txtdesc.SetFocus
                End If
            End Select
            
       End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       cmdLookup.Enabled = False
    ElseIf Range(Button.Index, 2, 3) Then
       cmdLookup.Enabled = True
    End If
    
    If PB_BlnkGls0 And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_SyProc, Me, txtCompCode, txtprocess, PR_Company, "BranchCnt", 3, "BranchCode", "BranchDesc", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
PB_BlnkGls0 = False
gc_dbcon.BeginTrans
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into sysBranch(Compcode,BranchCode,BranchDesc) VALUES ('" & txtCompCode & "','" & txtprocess.Text & "','" & txtdesc.Text & "')"
              cntsql.Execute
              
              txtprocess = DoPad(Val(0 & PR_Company.Fields("BranchCnt")) + 1, 3)
           Case "E"
              cntsql.CommandText = "UPDATE sysBranch SET BranchDesc= '" & txtdesc.Text & "' WHERE  CompCode = '" & txtCompCode & "' And BranchCode = '" & txtprocess.Text & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM  sysBranch WHERE CompCode = '" & txtCompCode & "' And BranchCode = '" & txtprocess.Text & "'"
              cntsql.Execute

     End Select
gc_dbcon.CommitTrans
     PR_SyProc.Requery

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
     
End Sub
Private Sub SetVal()
     txtdesc = PR_SyProc("Branchdesc")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtCompCode) = 3 And Len(txtprocess.Text) = 3 And txtdesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  PR_Company.Requery
  PR_SyProc.Requery
End Sub

