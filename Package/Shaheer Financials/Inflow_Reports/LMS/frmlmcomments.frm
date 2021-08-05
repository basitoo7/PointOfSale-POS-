VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmlmcomments 
   Caption         =   "Comments"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   Icon            =   "frmlmcomments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   7065
      Begin VB.TextBox txtconvdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1845
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   915
         Width           =   3135
      End
      Begin VB.ComboBox txtconvtype 
         Height          =   315
         ItemData        =   "frmlmcomments.frx":030A
         Left            =   1215
         List            =   "frmlmcomments.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   915
         Width           =   600
      End
      Begin VB.TextBox txtcomments 
         Height          =   1650
         Left            =   1230
         MaxLength       =   65535
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1290
         Width           =   5625
      End
      Begin VB.TextBox txtCustomerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2805
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   4020
      End
      Begin VB.CommandButton cmdlookup1 
         Height          =   315
         Left            =   2445
         Picture         =   "frmlmcomments.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   330
      End
      Begin MSMask.MaskEdBox txtCustomerNo 
         Height          =   315
         Left            =   1215
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox txtTransCode 
         Height          =   315
         Left            =   1215
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   10
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
      Begin MSComCtl2.DTPicker txttransdate 
         Height          =   285
         Left            =   5625
         TabIndex        =   9
         Top             =   195
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   37293
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date :"
         Height          =   210
         Left            =   4230
         TabIndex        =   12
         Top             =   210
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comments :"
         Height          =   195
         Left            =   315
         TabIndex        =   10
         Top             =   1230
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reference Id :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   195
         Width           =   1020
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer :"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   545
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conv.Type:"
         Height          =   195
         Left            =   315
         TabIndex        =   1
         Top             =   895
         Width           =   825
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7065
      _ExtentX        =   12462
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
               Picture         =   "frmlmcomments.frx":0496
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmlmcomments.frx":08EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmlmcomments.frx":0D3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmlmcomments.frx":1192
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmlmcomments.frx":15E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmlmcomments.frx":1A3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmlmcomments.frx":218E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmlmcomments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_Blnklmcomm As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Customer As New Recordset
Dim PR_lmComments As New Recordset


Private Sub cmdLookup1_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustomerNo
    Set PO_DESC = txtCustomerName
    
    GoTop PR_Customer
    MyLookup.Caption = "Customers"
    MyLookup.FillGrid PR_Customer, "CustomerNo", "CustomerName", txtCustomerNo.MaxLength
    MyLookup.Show 1
    
    If Len(txtCustomerNo) > 0 Then txtCustomerNo_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Call DentMode(Mode, 4, PR_lmComments, Me, txtTransCode, txtCustomerNo, ParaCntr_Rs, "LmComtCode", 10, "Transcode", "InflowType", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
' Setting up Preveliges
  
  SetToolBar(1) = chkRights("LMCOMMT001")
  SetToolBar(2) = chkRights("LMCOMMT002")
  SetToolBar(3) = chkRights("LMCOMMT003")
  SetToolBar(4) = chkRights("LMCOMMT004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  PR_Customer.Open "Select Customer.* from Customer Where Compcode+BranchCode = '" & Gs_compcode + Gs_BranchCode & "' Order By CustomerNo", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_lmComments.Open "Select lm_comments.* from lm_comments", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  PB_Blnklmcomm = IIf(PR_lmComments.EOF, True, False)
  txttransdate = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Customer.Close
    PR_lmComments.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If PB_Blnklmcomm And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
       Cancel = True
    'ElseIf Button.Index = 5 Then
    '   Call setprint
    Else
        Mode = DentMode(Mode, Button.Index, PR_lmComments, Me, txtTransCode, txtCustomerNo, ParaCntr_Rs, "LmComtCode", 10, "Transcode", "InflowType", 0, False, Toolbar1)
    End If
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_Blnklmcomm = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into LM_comments(Compcode,Branchcode,transcode,customerno,convtype,transdate,comments) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtTransCode.Text & "','" & txtCustomerNo.Text & "','" & txtconvtype.Text & "','" & Format(txttransdate.Value, "YYYY/MM/DD") & "','" & txtcomments.Text & "')"
              cntsql.Execute
                If Val(0 & txtTransCode) < (ParaCntr_Rs.Fields("LMComtCode") - 1) Then
                ' Auto Increment in DentMode Will (-) them here
                 ParaCntr_Rs.Fields("LMComtCode") = ParaCntr_Rs.Fields("LMComtCode") - 1
                 ParaCntr_Rs.Update
                 txtTransCode = DoPad(LTrim(Str(ParaCntr_Rs.Fields("LMComtCode") + 1)), 10)
               Else
                     txtTransCode = DoPad(LTrim(Str(ParaCntr_Rs.Fields("LMComtCode") + 1)), 10)
               End If

           Case "E"
              cntsql.CommandText = "UPDATE LM_comments SET customerno= '" & txtCustomerNo.Text & "', convtype= '" & txtconvtype.Text & "',  transdate= '" & Format(txttransdate.Value, "YYYY/MM/DD") & "', comments= '" & txtcomments.Text & "' WHERE  Compcode+BranchCode+transcode= '" & Gs_compcode + Gs_BranchCode + txtTransCode.Text & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM LM_comments WHERE  Compcode+BranchCode+transcode= '" & Gs_compcode + Gs_BranchCode + txtTransCode.Text & "'"
              cntsql.Execute
           
     End Select
     PR_lmComments.Requery
End Sub
Private Sub SetVal()
     txtCustomerNo = PR_lmComments("customerno")
     txttransdate = PR_lmComments("transdate")
     txtconvtype = PR_lmComments("convtype") & ""
     txtcomments = PR_lmComments("comments") & ""
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtTransCode.Text) = txtTransCode.MaxLength And txtCustomerNo <> "" And txtconvtype <> "" Then
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
Private Sub txtCustomerNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

If KeyCode = vbKeyReturn And txtCustomerNo.Text <> "" Then
      txtCustomerNo = DoPad(txtCustomerNo, txtCustomerNo.MaxLength)
        If Not MySeek(txtCustomerNo.Text, "CustomerNo", PR_Customer) Then
              Call SetErr(Gs_RecNFMsg, vbCritical)
              txtCustomerNo.SetFocus
        Else
             txtCustomerName.Text = PR_Customer("CustomerName")
              txtconvtype.SetFocus
        End If
   ElseIf KeyCode = vbKeyF12 Then
      Call cmdLookup1_Click
   ElseIf KeyCode = vbKeyPageUp Then
      txttransdate.SetFocus
   End If
End Sub

Private Sub txtTransCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Trim(txtTransCode.Text) <> "" Then
          
         txtTransCode.Text = DoPad(txtTransCode.Text, txtTransCode.MaxLength)
         lb_found = MySeek(txtTransCode.Text, "Transcode", PR_lmComments)
       
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                    SetClear Me
                   txtTransCode.SetFocus
                Else
                   txttransdate.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtTransCode.SetFocus
                Else
                   If Mode <> "D" Then txtCustomerNo.SetFocus
                   Call SetVal
                End If
            End Select
   End If
End Sub

Private Sub txttransdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCustomerNo.SetFocus
If KeyCode = vbKeyPageUp Then txtTransCode.SetFocus
End Sub
Public Sub FrmRefresh()
  PR_Customer.Requery
  PR_lmComments.Requery
End Sub
