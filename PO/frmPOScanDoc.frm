VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{50F16B18-467E-11D1-8271-00C04FC3183B}#1.0#0"; "shimgvw.dll"
Begin VB.Form frmPOScanDoc 
   Caption         =   "Scan Voucher Documents"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOScanDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   10980
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
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
            Caption         =   "&Slip"
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
               Picture         =   "frmPOScanDoc.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOScanDoc.frx":075E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOScanDoc.frx":0BB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOScanDoc.frx":1006
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOScanDoc.frx":145A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOScanDoc.frx":18AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOScanDoc.frx":2002
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7650
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   10935
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   570
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   615
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtLocCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   900
         Width           =   1470
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   4020
         Picture         =   "frmPOScanDoc.frx":2456
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   885
         Width           =   315
      End
      Begin VB.ComboBox txtnoteType 
         Height          =   330
         ItemData        =   "frmPOScanDoc.frx":25C8
         Left            =   2535
         List            =   "frmPOScanDoc.frx":25D2
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   180
         Width           =   3630
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   690
         Top             =   2850
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtscandoctitle 
         Height          =   345
         Left            =   2520
         MaxLength       =   255
         TabIndex        =   13
         Tag             =   "SKIP"
         Top             =   1275
         Width           =   4995
      End
      Begin VB.TextBox txtscanremarks 
         Height          =   510
         Left            =   2520
         MaxLength       =   255
         TabIndex        =   9
         Tag             =   "SKIP"
         Top             =   1680
         Width           =   8310
      End
      Begin VB.TextBox txtpath 
         Height          =   345
         Left            =   2520
         TabIndex        =   8
         Tag             =   "SKIP"
         Top             =   2280
         Width           =   7080
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   9660
         TabIndex        =   7
         Top             =   2265
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add To Grid"
         Height          =   390
         Left            =   8490
         TabIndex        =   3
         Top             =   2685
         Width           =   1170
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Clear Text"
         Height          =   390
         Left            =   9690
         TabIndex        =   2
         Top             =   2685
         Width           =   1170
      End
      Begin VB.Frame Frame2 
         Height          =   4530
         Left            =   60
         TabIndex        =   4
         Top             =   3045
         Width           =   10830
         Begin PREVIEWLibCtl.Preview Preview1 
            Height          =   4230
            Left            =   5175
            TabIndex        =   5
            Top             =   240
            Width           =   5580
         End
         Begin MSFlexGridLib.MSFlexGrid GrdGRN 
            Height          =   4245
            Left            =   75
            TabIndex        =   6
            Top             =   225
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   7488
            _Version        =   393216
            Rows            =   1
            BackColorFixed  =   -2147483637
         End
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   4740
         TabIndex        =   15
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63045633
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   2535
         TabIndex        =   16
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63045633
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker dtptransdate 
         Height          =   315
         Left            =   7155
         TabIndex        =   24
         Top             =   195
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63045633
         CurrentDate     =   37309
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Trans Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6240
         TabIndex        =   25
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Note # :"
         Height          =   210
         Left            =   1920
         TabIndex        =   22
         Top             =   915
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date :"
         Height          =   210
         Left            =   4065
         TabIndex        =   19
         Top             =   570
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1665
         TabIndex        =   18
         Top             =   585
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note Type :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1665
         TabIndex        =   17
         Top             =   210
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Scan Document Title :"
         Height          =   210
         Left            =   900
         TabIndex        =   12
         Top             =   1245
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Remarks for Scanning Document :"
         Height          =   210
         Left            =   60
         TabIndex        =   11
         Top             =   1680
         Width           =   2460
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Browse for File :"
         Height          =   210
         Left            =   1290
         TabIndex        =   10
         Top             =   2325
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPOScanDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGRN As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object


Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String

Dim PR_Dumy As New Recordset
Dim fs As New FileSystemObject
Dim ls_filename As String
Dim ls_fileExtension As String
Dim ls_filePath As String
Dim PR_Branch As New Recordset
Dim lb_found As Boolean

Private Sub Command1_Click()
CommonDialog1.ShowOpen
txtpath = CommonDialog1.FileName
ls_filename = CommonDialog1.FileTitle
ls_fileExtension = fs.GetExtensionName(ls_filename)

Preview1.ShowFile txtpath, 1
Preview1.Zoom 1

End Sub

Private Sub Command2_Click()
If txtLocCode = "" Then
    Call MsgBox("Enter/Select Note no!!!", vbCritical)
    txtLocCode.SetFocus
ElseIf txtscandoctitle = "" Then
    Call MsgBox("Enter Scan Document Title!!!", vbCritical)
    txtscandoctitle.SetFocus
ElseIf txtpath = "" Then
    Call MsgBox("Enter/Select File Path !!!", vbCritical)
    txtpath.SetFocus
Else
    Call AddToGrid
End If
End Sub

Private Sub Command3_Click()
txtscandoctitle = ""
txtscanremarks = ""
txtpath = ""
End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = Text1
    
   If txtnoteType.Text = "Demand Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_DemandNote "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Purchase Order" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POOrderNote "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Gate Pass Inward" Then
        Gs_SQL = "Select TransCode, TransDate from PO_GatePass "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Inspection Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_Inspection "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Inspection Return Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_InspectionReturn "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Good Receive Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POGRN "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POGRNReturn "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    
    End If
        MyLookupOLDB.Show 1
    
   If txtLocCode <> "" Then Call txtLocCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub dtptransdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtvno.SetFocus
End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Len(txtLocCode.Text) > 0 Then
    If txtnoteType.Text = "" Then
    Call MsgBox("Select Note Type !!!", vbCritical)
    txtnoteType.SetFocus
    Exit Sub
    End If
 End If
    
 If KeyCode = vbKeyReturn And Len(txtLocCode.Text) > 0 Then
 txtLocCode = DoPad(txtLocCode, txtLocCode.MaxLength)

 If PR_Dumy.State = 1 Then PR_Dumy.Close
    If txtnoteType.Text = "Demand Note" Then
        ls_sql = "Select TransCode, TransDate from PO_DemandNote "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Purchase Order" Then
        ls_sql = "Select TransCode, TransDate from PO_POOrderNote "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Gate Pass Inward" Then
        ls_sql = "Select TransCode, TransDate from PO_GatePass "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Inspection Note" Then
        ls_sql = "Select TransCode, TransDate from PO_Inspection "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Inspection Return Note" Then
        ls_sql = "Select TransCode, TransDate from PO_Inspectionreturn "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Good Receive Note" Then
        ls_sql = "Select TransCode, TransDate from PO_POGRN "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        ls_sql = "Select TransCode, TransDate from PO_POGRNReturn "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    End If
    PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
    If PR_Dumy.EOF Then
        Call MsgBox(txtnoteType.Text & " not found !!!", vbCritical)
        txtLocCode.SetFocus
    Else
        Text1.Text = PR_Dumy("TransDate")
        LoadGRNTrans
        txtscandoctitle.SetFocus
    End If

 ElseIf KeyCode = vbKeyF12 Then
        Command2_Click
 End If

End Sub

Private Sub txtpath_Change()
If txtpath <> "" Then
  Preview1.ShowFile txtpath, 1
  Preview1.Zoom 1
End If
End Sub

Private Sub txtscandoctitle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtscanremarks.SetFocus
End Sub

Private Sub txtscanremarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtpath.SetFocus
End Sub


Private Sub Form_Load()

 
  SetToolBar(1) = chkRights("ICISUSTP01")
  SetToolBar(2) = chkRights("ICISUSTP02")
  SetToolBar(3) = chkRights("ICISUSTP03")
  SetToolBar(4) = chkRights("ICISUSTP04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  InitializeGrid

  
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       InitializeGrid
    End If
    
    If Button.Index = 7 Then
    InitializeGrid
    End If
    
    If PB_BlnkGRN And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_ICIssue, Me, txtnoteType, txtnoteType, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
    End If
  
End Sub
Public Sub SaveValues()
'On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String
Dim ls_Transtype As String
Dim ls_Pmttranscode As String

gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
              gc_dbcon.Execute "DELETE FROM PO_ScanDocuments WHERE CompCode = '" & Gs_compcode & "' AND Notetype = " & txtnoteType.ListIndex & " and NoteNo = '" & txtLocCode & "' "
              
           Case Else
                
                gc_dbcon.Execute "DELETE FROM PO_ScanDocuments WHERE CompCode = '" & Gs_compcode & "' AND Notetype = " & txtnoteType.ListIndex & " and NoteNo = '" & txtLocCode & "' "
              
                
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      ls_filename = Trim(str(txtnoteType.ListIndex)) + "-" + txtLocCode + "." + .TextMatrix(ln_cnt, 6)
                      
                      fs.CopyFile .TextMatrix(ln_cnt, 5), Gs_POScanDoc & "\" & ls_filename
                      
                      ls_filePath = Gs_POScanDoc & "\" & ls_filename
                      
                      
                      ls_sql = "INSERT into PO_ScanDocuments(Compcode, NoteType,Noteno,TransDate, Title,Remarks ,FilePath)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "'," & txtnoteType.ListIndex & ",'" & Trim(.TextMatrix(ln_cnt, 3)) & "','" & Format(.TextMatrix(ln_cnt, 4), "YYYY/MM/DD") & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 2)) & "','" & ls_filePath & "')"
                      gc_dbcon.Execute ls_sql
                    Next
               End With
               
                 
                 
     End Select

gc_dbcon.CommitTrans
InitializeGrid

Exit Sub
RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Public Sub ClearVal()
End Sub
Private Sub setprint()
End Sub
Private Sub SetVal()
    
End Sub
Public Function ChkInputs() As Boolean
    If PI_SrNo > 0 Then
            ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function
Public Sub FrmRefresh()
End Sub
Private Sub AddToGrid()
Dim ln_cnt As Integer
       If txtLocCode <> "" Then
                    If PS_RowClicked = "" Then
                        If PI_SrNo = 0 Then
                            PI_SrNo = 1
                        Else
                            PI_SrNo = PI_SrNo + 1
                         End If
                     End If
        
                        With GrdGRN
                            If PS_RowClicked = "" Then
                                    If Not PI_SrNo = 1 Then .Rows = .Rows + 1
                                        .Row = .Rows - 1
                                    Else
                                        .Row = PI_CurRow
                                    End If
                                    If PS_RowClicked = "" Then
                                        .TextMatrix(.Row, 0) = PI_SrNo
                                    Else
                                        .TextMatrix(.Row, 0) = PI_CurRow
                                    End If
                                                .TextMatrix(.Row, 1) = txtscandoctitle
                                                .TextMatrix(.Row, 2) = txtscanremarks
                                                .TextMatrix(.Row, 3) = Trim(txtLocCode)
                                                .TextMatrix(.Row, 4) = dtptransdate
                                                .TextMatrix(.Row, 5) = txtpath
                                                .TextMatrix(.Row, 6) = ls_fileExtension
                                                
                                                
                                                
                               
                                txtscandoctitle = ""
                                txtscanremarks = ""
                                txtpath = ""
                                PS_RowClicked = ""
                                txtscandoctitle.SetFocus
                      
                           
                          End With
                     
        
       End If
      

End Sub
Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Scan Document Title|<Remarks|<Note No|<Note Date|<File Path|<FileExt"
        .ColWidth(1) = 2500
        .ColWidth(2) = 2000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 0
        .Redraw = True
    End With
    PI_SrNo = 0
    PI_CurRow = 0
    PS_RowClicked = ""
    Preview1.ShowFile "", 1
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
    With GrdGRN
        If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row

            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
            End If
        End If
    End With
End Sub
Private Sub LoadGRNTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
    
Pr_LoadTrans.Open "Select * from PO_ScanDocuments where Compcode = '" & Gs_compcode & "' and notetype= " & txtnoteType.ListIndex & " and Noteno = '" & txtLocCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("Title") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Remarks") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("Noteno") & "")
                .TextMatrix(.Row, 4) = Pr_LoadTrans("TransDate")
                .TextMatrix(.Row, 5) = Trim(Pr_LoadTrans("FilePath") & "")
                .TextMatrix(.Row, 6) = fs.GetExtensionName(.TextMatrix(.Row, 5))
                
                
                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
    End If
    txtscandoctitle.SetFocus
    Pr_LoadTrans.Close
    
Exit Sub
LocalErr:
Call MsgBox(Err.Description)
End Sub
Private Sub GrdGRN_DblClick()
    With GrdGRN
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
        If .TextMatrix(.Row, 1) <> "" Then
        txtscandoctitle = .TextMatrix(.Row, 1)
        txtscanremarks = .TextMatrix(.Row, 2)
        dtptransdate = .TextMatrix(.Row, 4)
        txtLocCode = .TextMatrix(.Row, 3)
        txtpath = .TextMatrix(.Row, 5)
        PS_RowClicked = "Y"
        txtscandoctitle.SetFocus
        End If
    End With
End Sub

Public Sub SetFrmEnv(ls_mode As String)
    txtLocCode.Enabled = IIf(ls_mode <> "D", True, False)
    txtpartycode.Enabled = IIf(ls_mode <> "D", True, False)
    txtremarks.Enabled = IIf(ls_mode <> "D", True, False)
    Frame2.Enabled = IIf(ls_mode <> "D", True, False)
End Sub

