VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOPostPurchase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post GRN"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
   Icon            =   "frmPOPostGRN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   15
      TabIndex        =   3
      Top             =   -15
      Width           =   5175
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1365
         TabIndex        =   0
         Top             =   435
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64880641
         CurrentDate     =   40949
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Preview"
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   2070
         Width           =   1110
      End
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2820
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1185
         Width           =   2190
      End
      Begin VB.TextBox txtTransNo 
         BackColor       =   &H00FFFF00&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "SKIPN"
         Top             =   1185
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2490
         Picture         =   "frmPOPostGRN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1185
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Post GRN"
         Height          =   315
         Left            =   3420
         TabIndex        =   6
         Top             =   2085
         Width           =   1710
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   30
         TabIndex        =   4
         Top             =   1530
         Width           =   5100
         Begin VB.Label lblStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   75
            TabIndex        =   7
            Top             =   180
            Width           =   4185
         End
      End
      Begin Crystal.CrystalReport rptVoucher 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   3
         WindowControlBox=   0   'False
         WindowMaxButton =   0   'False
         WindowMinButton =   0   'False
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1365
         TabIndex        =   1
         Top             =   810
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64880641
         CurrentDate     =   40949
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Date From  :"
         Height          =   255
         Left            =   75
         TabIndex        =   13
         Top             =   450
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Date To  :"
         Height          =   255
         Left            =   75
         TabIndex        =   12
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "GRN Note #  :"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   1215
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Post GRN "
         Height          =   450
         Left            =   60
         TabIndex        =   5
         Top             =   210
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmPOPostPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ls_transcode As String
Dim ls_transcodePS As String
Dim ls_sql  As String
Dim PR_Dumy As New Recordset
Dim pr_dumy1 As New Recordset
Dim ln_cnt
Dim PR_ICIssue As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object

Private Function maxtranscode() As String
pr_dumy1.Open "select max(transcode) as transcode from IC_TransMaster where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy1("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy1.Close
End Function
Private Function maxtranscodePS(ls_transtype As String) As String
pr_dumy1.Open "select max(InvoiceNo) as transcode from IC_TransMaster where compcode = '" & Gs_compcode & "' and transtype = '" & ls_transtype & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
    maxtranscodePS = DoPad(Trim(str(Int(0 & pr_dumy1("transcode")) + 1)), 10)
Else
    maxtranscodePS = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy1.Close
End Function
Private Sub Command1_Click()
Dim ls_res

If Trim(txttransno) = "" Then
Call MsgBox("Enter/Select " & Label3.Caption & " !!!", vbExclamation)
Exit Sub
End If

If Command1.Caption = "Post GRN" Then
    lblStatus = "Posting of GRN in Progress..."
    DoEvents
    If PostPurchaseVoucher(Gs_compcode, DTPicker1, DTPicker2, txttransno) Then
        Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
    End If
   
    lblStatus = ""

ElseIf Command1.Caption = "Post GRRN" Then
    lblStatus = "Posting of GRRN in Progress..."
    DoEvents
    If PostPurchaseReturnVoucher(Gs_compcode, DTPicker1, DTPicker2, txttransno) Then
      Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
    End If
    lblStatus = ""
ElseIf Command1.Caption = "Post Payments" Then
    lblStatus = "Posting of Payments in Progress..."
    DoEvents
    Call PostPayableVoucherBank(Gs_compcode, DTPicker1, DTPicker2, txttransno)
    Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
    lblStatus = ""
        
ElseIf Command1.Caption = "Post Payments Cash" Then
    lblStatus = "Posting of Payments in Progress..."
    DoEvents
    Call PostPayableVoucherCash(Gs_compcode, DTPicker1, DTPicker2, txttransno)
    Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
    lblStatus = ""
End If
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttransno
    Set PO_DESC = Text1
    If Label1.Caption = "Post GRN" Then
        Gs_SQL = "SELECT GRN.TransCode AS ComputerCode, GRN.GRNCode AS GRNCode, Vendors.Description AS 'Vendors.Description', GRN.TransDate AS GRNDate,    GRN.NetAmount AS 'GRN.NetAmount' FROM         PO_POGRN GRN INNER JOIN         IC_Supplier Vendors ON GRN.Compcode = Vendors.Compcode AND GRN.AccountCode = Vendors.SupplierCode"
        Gs_OrderBy = "ORDER BY GRN.TransCode desc"
        Gs_OtherPara = " Where GRN.compcode = '" & Gs_compcode & "'  and  grn.glstatus = 0"
        frmPosearchRecords.Caption = "GRN"
        frmPosearchRecords.Show 1
        If txttransno <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)
        Exit Sub
   ElseIf Label1.Caption = "Post Payable" Then
        Gs_SQL = "Select TransCode,Transdate from PO_PayableMaster "
        Gs_FindFld = "TransCode"
        Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "' and glstatus = 0"
        Gs_OrderBy = "Order by TransCode"
        MyLookupOLDB.Caption = "Vendor Payable"
    ElseIf Label1.Caption = "Post Payable Cash" Then
        
        Gs_SQL = "Select TransCode,Transdate from PO_PayablecashMaster "
        Gs_FindFld = "TransCode"
        Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "' and glstatus = 0 "
        Gs_OrderBy = "Order by TransCode"
        MyLookupOLDB.Caption = "Vendor Payable"
    
    ElseIf Label1.Caption = "Post GRRN" Then
        Gs_SQL = "Select TransCode,Transdate from PO_POGRNReturn "
        Gs_FindFld = "TransCode"
        Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "' and glstatus = 0"
        Gs_OrderBy = "Order by TransCode"
        MyLookupOLDB.Caption = "GRRN"
    End If
    MyLookupOLDB.Show 1
    
    If txttransno <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command3_Click()
If Label1.Caption = "Post GRN" Then
    PrintGRNnote
ElseIf Label1.Caption = "Post GRRN" Then
    PrintGRRNnote
End If
End Sub
Private Sub PrintGRNnote()
On Error GoTo LocalErr

   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POGRN.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Good Receive Note'"
        .SelectionFormula = "{PO_POOrderNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & DTPicker1.Year & "," & DTPicker1.Month & "," & DTPicker1.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & DTPicker2.Year & "," & DTPicker2.Month & "," & DTPicker2.Day & ") "
        If txttransno <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrderNote.transcode} = '" & Trim(txttransno) & "'"
        End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub
Private Sub PrintGRRNnote()
On Error GoTo LocalErr

   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POGRNReturn.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Good Receive Return Note'"
        .SelectionFormula = "{PO_POOrderNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & DTPicker1.Year & "," & DTPicker1.Month & "," & DTPicker1.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & DTPicker2.Year & "," & DTPicker2.Month & "," & DTPicker2.Day & ") "
        If txttransno <> "" Then
              .SelectionFormula = .SelectionFormula & "  and {PO_POOrderNote.transcode} = '" & Trim(txttransno) & "'"
        End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
DTPicker2.SetFocus
End If
End Sub
Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txttransno.SetFocus
End If
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn And Len(txttransno.Text) > 0 Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txttransno.Text = DoPad(UCase(txttransno.Text), 10)
         If Label1.Caption = "Post GRN" Then
            PR_ICIssue.Open "select * from PO_POGRN where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         ElseIf Label1.Caption = "Post Payable" Then
            PR_ICIssue.Open "select * from PO_PayableMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         ElseIf Label1.Caption = "Post Payable Cash" Then
            PR_ICIssue.Open "select * from PO_PayablecashMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         Else
            PR_ICIssue.Open "select * from PO_POGRNReturn where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         End If
         If PR_ICIssue.EOF Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txttransno.SetFocus
         Else
            Text1 = PR_ICIssue("TransDate")
         End If
  ElseIf KeyCode = vbKeyReturn And Len(txttransno.Text) = 0 Then
        Call Command2_Click
   End If
End Sub
