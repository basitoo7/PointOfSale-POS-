VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmLmDocument 
   Caption         =   "Lease Documents"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   Icon            =   "FrmLmDocument.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1980
      Left            =   15
      TabIndex        =   0
      Top             =   570
      Width           =   6570
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   6165
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1605
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2055
         Picture         =   "FrmLmDocument.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   165
         Width           =   315
      End
      Begin VB.TextBox lbldesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   165
         Width           =   4095
      End
      Begin VB.ComboBox txtdocstatus 
         Height          =   315
         ItemData        =   "FrmLmDocument.frx":047C
         Left            =   1470
         List            =   "FrmLmDocument.frx":0486
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1590
         Width           =   1320
      End
      Begin VB.TextBox TxtDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1470
         MaxLength       =   65535
         TabIndex        =   9
         Top             =   915
         Width           =   5025
      End
      Begin VB.CommandButton cmdlookup1 
         Height          =   315
         Left            =   2055
         Picture         =   "FrmLmDocument.frx":04A3
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtDocCode 
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   4
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
      Begin MSMask.MaskEdBox txtdoctype 
         Height          =   315
         Left            =   1470
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   570
         _ExtentX        =   1005
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   195
         Left            =   525
         TabIndex        =   8
         Top             =   885
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Document Type :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   195
         Width           =   1230
      End
      Begin VB.Label lblFieldLabel 
         Caption         =   "Document Code :"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   5
         Top             =   555
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status :"
         Height          =   195
         Left            =   870
         TabIndex        =   1
         Top             =   1605
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   1005
      ButtonWidth     =   1376
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
         Left            =   4560
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
               Picture         =   "FrmLmDocument.frx":0615
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDocument.frx":0A69
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDocument.frx":0EBD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDocument.frx":1311
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDocument.frx":1765
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDocument.frx":1BB9
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDocument.frx":230D
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmLmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLmDoc As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_DocType As New Recordset
Dim PR_Documets As New Recordset


Private Sub cmdLookup1_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtDocCode
    Set PO_DESC = Text1
    
    GoTop PR_Documets
    If PR_Documets.RecordCount > 0 Then
        PR_Documets.Filter = "Doctype = '" & txtdoctype & "'"
        MyLookup.Caption = "Documents"
        MyLookup.FillGrid PR_Documets, "DocCode", "Doctype", txtDocCode.MaxLength
        MyLookup.Show 1
        PR_Documets.Filter = adFilterNone
    Else
        Call SetErr(Gs_RecNFMsg, vbCritical)
    End If
    If Len(txtDocCode) > 0 Then txtDocCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtdoctype
    Set PO_DESC = lbldesc
    
    GoTop PR_DocType
    MyLookup.Caption = "Document Types"
    MyLookup.FillGrid PR_DocType, "Codeid", "Description", txtdoctype.MaxLength
    MyLookup.Show 1
    
    If Len(txtdoctype) > 0 Then txtdoctype_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Call DentMode(Mode, 4, PR_documents, Me, txtdoctype, txtDocCode, ParaCntr_Rs, "LmComtCode", 10, "Transcode", "InflowType", 1, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
' Setting up Preveliges
  
  SetToolBar(1) = chkRights("LMDOCUMEN1")
  SetToolBar(2) = chkRights("LMDOCUMEN1")
  SetToolBar(3) = chkRights("LMDOCUMEN1")
  SetToolBar(4) = chkRights("LMDOCUMEN1")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  PR_DocType.Open "Select CustClasses.* from CustClasses  Where CustClasses.Codestat = '1' order  by Codeid", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Documets.Open "Select lm_Documents.*,lm_Documents.DocType+lm_Documents.DocCode as FindFld  from lm_Documents  where Compcode+Branchcode = '" & Gs_compcode + Gs_BranchCode & "' order by Findfld ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  PB_BlnkLmDoc = IIf(PR_Documets.EOF, True, False)
  txttransdate = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_DocType.Close
    PR_Documets.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If PB_Blnklmcomm And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
       Cancel = True
    'ElseIf Button.Index = 5 Then
    '   Call setprint
    Else
        Mode = DentMode(Mode, Button.Index, PR_Documets, Me, txtdoctype, txtdoctype, ParaCntr_Rs, "LmComtCode", 10, "Transcode", "InflowType", 1, False, Toolbar1)
    End If
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkLmDoc = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into LM_Documents(Compcode,Branchcode,DocType,Doccode,Docdescrip,Docstatus) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtdoctype & "','" & txtDocCode & "','" & txtDesc & "','" & UCase(Left(txtdocstatus, 1)) & "')"
              cntsql.Execute
              'PR_DocType ("Codecntr")
              
           Case "E"
              cntsql.CommandText = "UPDATE LM_Documents SET Docdescrip= '" & txtDesc.Text & "', DocStatus = '" & txtdocstatus.Text & "' WHERE  Compcode+BranchCode+DocType+DocCode = '" & Gs_compcode + Gs_BranchCode + txtdoctype + txtDocCode & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM LM_Documents WHERE  Compcode+BranchCode+DocType+DocCode = '" & Gs_compcode + Gs_BranchCode + txtdoctype + txtDocCode & "'"
              cntsql.Execute
     End Select
     PR_Documets.Requery
End Sub
Private Sub SetVal()
     txtDesc = PR_Documets("DocDescrip")
     txtdocstatus = IIf(PR_Documets("DocStatus") = "M", "Must Fullfill", "Skipable")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtdoctype.Text) = txtdoctype.MaxLength And txtDocCode <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Private Sub txtconvtype_Click()
If txtconvtype.Text = "O" Then txtconvdesc.Text = "Office Meeting"
If txtconvtype.Text = "T" Then txtconvdesc.Text = "Telephone"
If txtconvtype.Text = "V" Then txtconvdesc.Text = "Visit"
End Sub

Private Sub txtconvtype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtcomments.SetFocus
If KeyCode = vbKeyPageUp Then txtconvtype.SetFocus
End Sub

Private Sub txtDocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Trim(txtDocCode.Text) <> "" Then
          
         txtDocCode.Text = DoPad(txtDocCode.Text, txtDocCode.MaxLength)
         lb_found = MySeek(txtDocCode.Text, "DocCode", PR_Documets)
       
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   txtDocCode.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txtDocCode.SetFocus
                Else
                   Call SetVal
                End If
            End Select
ElseIf KeyCode = vbKeyF12 Then
       cmdLookup1_Click
ElseIf KeyCode = vbKeyPageUp Then
       txtdoctype.SetFocus
End If
End If
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtdocstatus.SetFocus
If KeyCode = vbKeyPageUp Then txtDocCode.SetFocus
End Sub
Public Sub FrmRefresh()
  PR_Documets.Requery
  PR_DocType.Requery
End Sub

Private Sub txtdocstatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyPageUp Then txtDesc.SetFocus
End Sub

Private Sub txtdoctype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtdoctype <> "" Then
       txtdoctype = DoPad(txtdoctype, txtdoctype.MaxLength)
    If MySeek(txtdoctype, "Doctype", PR_DocType) Then
        lbldesc = PR_DocType("Description")
          If Mode = "A" Then
                txtDocCode = DoPad(PR_DocType("codecntr") + 1, txtDocCode.MaxLength)
                txtDocCode.Enabled = False
                cmdlookup1.Enabled = False
                txtDesc.SetFocus
          Else
                txtDocCode.Enabled = True
                cmdlookup1.Enabled = True
                txtDocCode.SetFocus
          End If

Else
     Call SetErr(Gs_RecNFMsg, vbCritical)
     txtdoctype.SetFocus
End If
ElseIf KeyCode = vbKeyF12 Then
       Command1_Click
End If
End Sub
