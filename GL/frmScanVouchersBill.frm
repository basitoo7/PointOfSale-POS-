VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScanVoucherBill 
   Caption         =   "Scan Voucher Documents"
   ClientHeight    =   8160
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
   Icon            =   "frmScanVouchersBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10980
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1058
      ButtonWidth     =   1402
      ButtonHeight    =   1005
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
               Picture         =   "frmScanVouchersBill.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmScanVouchersBill.frx":075E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmScanVouchersBill.frx":0BB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmScanVouchersBill.frx":1006
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmScanVouchersBill.frx":145A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmScanVouchersBill.frx":18AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmScanVouchersBill.frx":2002
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7560
      Left            =   30
      TabIndex        =   1
      Top             =   525
      Width           =   10935
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   690
         Top             =   2850
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtscandoctitle 
         Height          =   345
         Left            =   2535
         MaxLength       =   255
         TabIndex        =   20
         Top             =   720
         Width           =   4995
      End
      Begin VB.TextBox txtvchrtype 
         BackColor       =   &H00C0C000&
         Height          =   345
         Left            =   2550
         TabIndex        =   19
         Top             =   270
         Width           =   750
      End
      Begin VB.TextBox txtvchrdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   3675
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   12
         Tag             =   "SKIP"
         Top             =   270
         Width           =   2385
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
         Left            =   3315
         Picture         =   "frmScanVouchersBill.frx":2456
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtscanremarks 
         Height          =   870
         Left            =   2535
         MaxLength       =   255
         TabIndex        =   10
         Top             =   1140
         Width           =   8310
      End
      Begin VB.TextBox txtpath 
         Height          =   345
         Left            =   2535
         TabIndex        =   9
         Top             =   2085
         Width           =   7080
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   9660
         TabIndex        =   8
         Top             =   2070
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add To Grid"
         Height          =   390
         Left            =   8490
         TabIndex        =   5
         Top             =   2685
         Width           =   1170
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Clear Text"
         Height          =   390
         Left            =   9690
         TabIndex        =   4
         Top             =   2685
         Width           =   1170
      End
      Begin VB.TextBox txtvno 
         BackColor       =   &H00C0C000&
         Height          =   345
         Left            =   9765
         MaxLength       =   10
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtptransdate 
         Height          =   315
         Left            =   7245
         TabIndex        =   3
         Top             =   270
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12632064
         Format          =   63438849
         CurrentDate     =   40288
      End
      Begin VB.Frame Frame2 
         Height          =   4530
         Left            =   60
         TabIndex        =   6
         Top             =   2985
         Width           =   10830
         Begin MSFlexGridLib.MSFlexGrid GrdGRN 
            Height          =   4245
            Left            =   75
            TabIndex        =   7
            Top             =   225
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   7488
            _Version        =   393216
            Rows            =   1
            BackColorFixed  =   -2147483637
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Voucher Type :"
         Height          =   210
         Left            =   1380
         TabIndex        =   18
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Scan Document Title :"
         Height          =   210
         Left            =   945
         TabIndex        =   17
         Top             =   750
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Remarks for Scanning Document :"
         Height          =   210
         Left            =   90
         TabIndex        =   16
         Top             =   1155
         Width           =   2460
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Browse for File :"
         Height          =   210
         Left            =   1320
         TabIndex        =   15
         Top             =   2130
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Voucher Date :"
         Height          =   210
         Left            =   6105
         TabIndex        =   14
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Voucher No :"
         Height          =   210
         Left            =   8730
         TabIndex        =   13
         Top             =   300
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmScanVoucherBill"
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

Dim pr_dumy As New Recordset
Dim fs As New FileSystemObject
Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    
    Set PO_CODE = txtvchrType
    Set PO_DESC = txtVchrDesc
    Gs_SQL = "Select VchrType, VchrDescrip from Glvchrtype"
    Gs_FindFld = "VchrDescrip"
    Gs_OrderBy = "Order by VchrDescrip"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Voucher Type"
    MyLookupOLDB.Show 1
    
    If txtvchrType <> "" Then Call txtVchrType_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
txtpath = CommonDialog1.FileName
Preview1.ShowFile txtpath, 1
Preview1.Zoom 1

End Sub

Private Sub Command2_Click()
If txtvchrType = "" Then
    Call MsgBox("Enter/Select Voucher Type!!!", vbCritical)
    txtvchrType.SetFocus
ElseIf txtvno = "" Then
    Call MsgBox("Enter Voucher No!!!", vbCritical)
    txtvno.SetFocus
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
txtVchrDesc = ""
txtvchrType = ""
txtvno = ""
txtscandoctitle = ""
txtscanremarks = ""
txtpath = ""
End Sub

Private Sub dtptransdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtvno.SetFocus
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

Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtvchrType) <> "" And KeyCode = vbKeyReturn Then
        
        pr_dumy.Open "Select VchrType, VchrDescrip from Glvchrtype where Compcode  = '" & Gs_compcode & "' and vchrtype = '" & txtvchrType & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Voucher Type not found !!!", vbCritical)
            txtvchrType = ""
            txtVchrDesc = ""
            txtvchrType.SetFocus
        Else
            txtVchrDesc = pr_dumy("VchrDescrip")
            dtptransdate.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(txtvchrType) = "" And KeyCode = vbKeyReturn Then
        txtvchrType = ""
        txtVchrDesc = ""
End If
End Sub
Private Sub txtvno_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtvno) <> "" And KeyCode = vbKeyReturn Then
        txtvno = DoPad(txtvno, txtvno.MaxLength)
        pr_dumy.Open "Select Voucher_no from Gl_Ref where Compcode  = '" & Gs_compcode & "' and vchrtype = '" & txtvchrType & "' and Voucher_no = '" & txtvno & "' and Value_Date = '" & Format(dtptransdate, "YYYY/MM/DD") & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Voucher No. not found !!!", vbCritical)
            txtvno = ""
            txtvno.SetFocus
        Else
              Call LoadGRNTrans
        End If
        pr_dumy.Close

ElseIf Trim(txtvno) = "" And KeyCode = vbKeyReturn Then
        txtvno = ""
End If

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
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_ICIssue, Me, txtvchrType, txtvchrType, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
    End If
  
End Sub
Public Sub SaveValues()
'On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String
Dim ls_transtype As String
Dim ls_Pmttranscode As String

gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
              gc_dbcon.Execute "DELETE FROM Gl_ScanDocuments WHERE CompCode = '" & Gs_compcode & "' AND Value_Date = '" & Format(dtptransdate, "YYYY/MM/DD") & "' and vchrtype = '" & txtvchrType & "'  and Voucher_no = '" & txtvno & "'"
              
           Case Else
                
                gc_dbcon.Execute "DELETE FROM Gl_ScanDocuments WHERE CompCode = '" & Gs_compcode & "' AND Value_Date = '" & Format(dtptransdate, "YYYY/MM/DD") & "' and vchrtype = '" & txtvchrType & "'  and Voucher_no = '" & txtvno & "'"
              
                
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      ls_sql = "INSERT into Gl_ScanDocuments(Compcode, BranchCode,  Title,Remarks, VchrType,Value_Date, Trans_Date, Voucher_no,FilePath)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 2)) & "','" & Trim(.TextMatrix(ln_cnt, 3)) & "','" & Format(.TextMatrix(ln_cnt, 4), "YYYY/MM/DD") & "','" & Format(Date, "YYYY/MM/DD") & "','" & .TextMatrix(ln_cnt, 5) & "','" & .TextMatrix(ln_cnt, 6) & "')"
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
       If txtvchrType <> "" Then
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
                                                .TextMatrix(.Row, 3) = Trim(txtvchrType)
                                                .TextMatrix(.Row, 4) = dtptransdate
                                                .TextMatrix(.Row, 5) = txtvno
                                                
                                                'If Not fs.FolderExists(App.Path & Gs_GLScanDoc & "\" & txtvchrtype) Then
                                                'a = fs.BuildPath(App.Path & Gs_GLScanDoc, txtvchrtype)
                                                   
                                                '    fs.CreateFolder (App.Path & Gs_GLScanDoc & "\" & txtvchrtype)
                                                'End If
                                                
                                              '  Call fs.CopyFile(txtpath, App.Path & "\" & Gs_GLScanDoc)
                                                .TextMatrix(.Row, 6) = txtpath
                                                
                                                
                               
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
        .FormatString = "Sr# |<Scan Document Title|<Remarks|<V.Type|<V.Date|<V.#|<File Path"
        .ColWidth(1) = 2500
        .ColWidth(2) = 2000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        
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
    
Pr_LoadTrans.Open "Select * from Gl_ScanDocuments where Compcode = '" & Gs_compcode & "' and VchrType = '" & txtvchrType & "' and Voucher_no = '" & txtvno & "' and Value_Date = '" & Format(dtptransdate, "YYYY/MM/DD") & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("Title") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Remarks") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("vchrtype") & "")
                .TextMatrix(.Row, 4) = Pr_LoadTrans("Value_Date")
                .TextMatrix(.Row, 5) = Trim(Pr_LoadTrans("Voucher_no") & "")
                .TextMatrix(.Row, 6) = Trim(Pr_LoadTrans("FilePath") & "")
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
        txtvchrType = .TextMatrix(.Row, 3)
        If Len(txtvchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift
        dtptransdate = .TextMatrix(.Row, 4)
        txtvno = .TextMatrix(.Row, 5)
        txtpath = .TextMatrix(.Row, 6)
        PS_RowClicked = "Y"
        txtscandoctitle.SetFocus
        End If
    End With
End Sub

Public Sub SetFrmEnv(ls_mode As String)
    txtLocCode.Enabled = IIf(ls_mode <> "D", True, False)
    txtpartycode.Enabled = IIf(ls_mode <> "D", True, False)
    TxtRemarks.Enabled = IIf(ls_mode <> "D", True, False)
    Frame2.Enabled = IIf(ls_mode <> "D", True, False)
End Sub

