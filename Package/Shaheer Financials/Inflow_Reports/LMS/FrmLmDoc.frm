VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLmDoc 
   Caption         =   "Documents"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   Icon            =   "FrmLmDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Tag             =   " v      "
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   15
      TabIndex        =   3
      Top             =   585
      Width           =   6570
      Begin MSComCtl2.DTPicker DtpTDate 
         Height          =   315
         Left            =   5295
         TabIndex        =   20
         Top             =   1275
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57016321
         CurrentDate     =   37511
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2235
         Picture         =   "FrmLmDoc.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   555
         Width           =   315
      End
      Begin VB.TextBox txtleaseno 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   562
         Width           =   810
      End
      Begin VB.CommandButton cmdLookup 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2235
         Picture         =   "FrmLmDoc.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Width           =   315
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2565
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   195
         Width           =   3870
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   4935
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   555
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2220
         Picture         =   "FrmLmDoc.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   915
         Width           =   315
      End
      Begin VB.TextBox lbldesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2565
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   915
         Width           =   3885
      End
      Begin VB.ComboBox txtdocstatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmLmDoc.frx":0760
         Left            =   3165
         List            =   "FrmLmDoc.frx":076A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1275
         Width           =   1155
      End
      Begin VB.CommandButton cmdlookup1 
         Height          =   315
         Left            =   2220
         Picture         =   "FrmLmDoc.frx":0788
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1260
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtDocCode 
         Height          =   315
         Left            =   1425
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1275
         Width           =   810
         _ExtentX        =   1429
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
         Left            =   1425
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   915
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777088
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
      Begin MSMask.MaskEdBox txtCustNO 
         Height          =   315
         Left            =   1425
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   195
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777088
         PromptInclude   =   0   'False
         MaxLength       =   6
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Target Date :"
         Height          =   195
         Index           =   2
         Left            =   4335
         TabIndex        =   19
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lease # :"
         Height          =   195
         Index           =   0
         Left            =   705
         TabIndex        =   18
         Top             =   597
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer Code :"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Document Type :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   954
         Width           =   1230
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Document Code :"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   1305
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status :"
         Height          =   195
         Index           =   1
         Left            =   2580
         TabIndex        =   4
         Top             =   1305
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
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
               Picture         =   "FrmLmDoc.frx":08FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDoc.frx":0D4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDoc.frx":11A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDoc.frx":15F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDoc.frx":1A4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDoc.frx":1E9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLmDoc.frx":25F2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmLmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLmDoc As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_LMSInfo As New Recordset
Dim PR_Customer As New Recordset
Dim PR_DocType As New Recordset
Dim PR_Documets As New Recordset
Dim PR_LmDocumets As New Recordset
Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustNO
    Set PO_DESC = Text4
    Gs_SQL = "Select Customer.Customerno 'Customer No', Customer.CustomerName  'Customer Name' from Customer Inner Join Facilities On Customer.CustomerNo = Facilities.CustomerNo"
    Gs_FindFld = "CustomerName"
    Gs_OtherPara = " Where Customer.Compcode+Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo = '01'"
    Gs_OrderBy = "Order by Customer.CustomerNo,Customer.CustomerName"
    MyLookupOLDB.Caption = "Customers"
    MyLookupOLDB.Show 1
    
    If Len(txtCustNO) > 0 Then TxtCustNo_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtleaseno
    Set PO_DESC = Text1

    PR_LMSInfo.Filter = "BranchCode = '" & Gs_BranchCode & "' And CustomerNo = '" & txtCustNO & "'"
    GoTop PR_LMSInfo
    MyLookup.Caption = "Lease Agreements"
    MyLookup.FillGrid PR_LMSInfo, "LeaseNo", "LeaseAmount", txtleaseno.MaxLength
    MyLookup.Show 1
    If Len(txtleaseno) > 0 Then txtleaseno_KeyDown vbKeyReturn, vbKeyShift
    PR_LMSInfo.Filter = adFilterNone
End Sub
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
  
  SetToolBar(1) = chkRights("LMDOCUME01")
  SetToolBar(2) = chkRights("LMDOCUME02")
  SetToolBar(3) = chkRights("LMDOCUME03")
  SetToolBar(4) = chkRights("LMDOCUME04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  PR_Customer.Open "Select Customer.* from Customer Inner Join Facilities On Customer.CustomerNo = Facilities.CustomerNo Where Customer.Compcode+Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo = '01' Order By Customer.CustomerNo,Customer.BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_LMSInfo.Open "Select *,CustomerNo+LeaseNo As FindFld from LM_LeaseInfo where Compcode+Branchcode='" & Gs_compcode + Gs_BranchCode & "' Order by BranchCode,CustomerNo,LeaseNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_DocType.Open "Select CustClasses.* from CustClasses  Where CustClasses.Codestat = '1' order  by Codeid", gc_dbcon, adOpenStatic, adLockPessimistic, 1
  PR_Documets.Open "Select lm_Documents.*,lm_Documents.DocType+lm_Documents.DocCode as FindFld  from lm_Documents  where Compcode+Branchcode = '" & Gs_compcode + Gs_BranchCode & "' and FacilityNo = '01' order by Findfld ", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
  PR_LmDocumets.Open "Select Lm_Docstatus.*,Lm_Docstatus.CustomerNo+Lm_Docstatus.LeaseNo+Lm_Docstatus.Doctype+Lm_Docstatus.DocCode as Findfld from Lm_Docstatus where compcode+BranchCode = '" & Gs_compcode + Gs_BranchCode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PB_BlnkLmDoc = IIf(PR_LmDocumets.EOF, True, False)
  txttransdate = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Customer.Close
    PR_LMSInfo.Close
    PR_DocType.Close
    PR_Documets.Close
    PR_LmDocumets.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If PB_Blnklmcomm And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
       Cancel = True
    'ElseIf Button.Index = 5 Then
    '   Call setprint
    Else
        Mode = DentMode(Mode, Button.Index, PR_LmDocumets, Me, txtCustNO, txtCustNO, ParaCntr_Rs, "LmComtCode", 10, "Transcode", "InflowType", 1, False, Toolbar1)
    End If
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkLmDoc = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into LM_DocStatus(Compcode,Branchcode,Customerno ,leaseNo , DocType,Doccode,Docavailable,targetdate) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtCustNO) & "','" & Trim(txtleaseno) & "','" & Trim(txtdoctype) & "','" & Trim(txtDocCode) & "','" & Trim(UCase(Left(txtdocstatus, 1))) & "','" & Format(DtpTDate, "YYYY/MM/DD") & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE LM_DocStatus SET DocType= '" & txtdoctype.Text & "', DocCode = '" & txtDocCode.Text & "', Docabailable = '" & UCase(Left(txtdocstatus, 1)) & "' WHERE  Compcode+BranchCode+CustomerNo+LeaseNo+Doctype+Doccode = '" & Gs_compcode + Gs_BranchCode + Trim(txtCustNO) + Trim(txtleaseno) + Trim(txtdoctype) + Trim(txtDocCode) & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM LM_DocStatus  WHERE  Compcode+BranchCode+CustomerNo+LeaseNo+Doctype+Doccode = '" & Gs_compcode + Gs_BranchCode + Trim(txtCustNO) + Trim(txtleaseno) + Trim(txtdoctype) + Trim(txtDocCode) & "'"
              cntsql.Execute
     End Select
     PR_Documets.Requery
     PR_LmDocumets.Requery
     PR_DocType.Filter = adFilterNone
     PR_DocType.Requery
End Sub
Private Sub SetVal()
     txtdocstatus = IIf(PR_LmDocumets("DocAvailable") = "A", "Available", "Not Available")
     DtpTDate = PR_LmDocumets("TargetDate")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtdoctype.Text) = txtdoctype.MaxLength And txtDocCode <> "" And txtCustNO <> "" And txtleaseno <> "" Then
          ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function
Private Sub TxtCustNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If Lastkey(KeyCode) And Trim(txtCustNO) <> "" Then
   txtCustNO = DoPad(txtCustNO, txtCustNO.MaxLength)
   If MySeek(txtCustNO, "CustomerNo", PR_Customer) Then
      Text4 = PR_Customer("CustomerName") & ""
            txtleaseno.SetFocus
   Else
      Call SetErr(Gs_RecNFMsg, vbCritical)
      txtCustNO.SetFocus
   End If
 ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
 End If
End Sub

Private Sub txtDocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If KeyCode = vbKeyReturn And Trim(txtDocCode.Text) <> "" Then
          txtDocCode.Text = DoPad(txtDocCode.Text, txtDocCode.MaxLength)
          Select Case Mode
            Case "A"
                lb_found = MySeek(Trim(txtDocCode), "DocCode", PR_Documets)
                If Not lb_found Then
                     Call SetErr(Gs_RecFdMsg, vbCritical)
                     txtDocCode.SetFocus
                Else
                   If Not MySeek(Trim(txtCustNO) + Trim(txtleaseno) + Trim(txtdoctype) + Trim(txtDocCode), "Findfld", PR_LmDocumets) Then
                     txtdocstatus.SetFocus
                   Else
                     Call SetErr(Gs_RecFdMsg, vbCritical)
                     txtDocCode.SetFocus
                   End If
                End If
        Case Else
            lb_found = MySeek(Trim(txtCustNO) + Trim(txtleaseno) + Trim(txtdoctype) + Trim(txtDocCode), "Findfld", PR_LmDocumets)
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
End Sub

Public Sub FrmRefresh()
  PR_Customer.Requery
  PR_LMSInfo.Requery
  PR_Documets.Requery
  PR_DocType.Requery
End Sub

Private Sub txtdocstatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Trim(UCase(Left(txtdocstatus, 1))) = "N" Then
            DtpTDate.Enabled = True
            DtpTDate.SetFocus
    Else
            DtpTDate.Enabled = False
    End If
ElseIf KeyCode = vbKeyPageUp Then
    txtdesc.SetFocus
End If
End Sub

Private Sub txtdoctype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtdoctype <> "" Then
       txtdoctype = DoPad(txtdoctype, txtdoctype.MaxLength)
       If MySeek(txtdoctype, "CodeId", PR_DocType) Then
            lbldesc = PR_DocType("Description")
            txtDocCode.SetFocus
       Else
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtdoctype.SetFocus
End If
ElseIf KeyCode = vbKey Then
ElseIf KeyCode = vbKeyF12 Then
       Command1_Click
End If
End Sub
Private Sub txtleaseno_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If Lastkey(KeyCode) And Trim(txtleaseno) <> "" Then
   txtleaseno = DoPad(txtleaseno, txtleaseno.MaxLength)
    If MySeek(txtCustNO + txtleaseno, "FindFld", PR_LMSInfo) Then
        txtdoctype.SetFocus
    Else
        Call SetErr(Gs_RecNFMsg, vbCritical)
        txtleaseno.SetFocus
    End If
ElseIf KeyCode = vbKeyF12 Then
        Call Command1_Click
ElseIf KeyCode = vbKeyPageUp Then
        txtCustNO.SetFocus
End If
End Sub

