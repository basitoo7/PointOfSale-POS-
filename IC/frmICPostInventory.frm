VERSION 5.00
Begin VB.Form frmICPostInvntory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post Inventory Consumption Voucher"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   Icon            =   "frmICPostInventory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   15
      TabIndex        =   1
      Top             =   -15
      Width           =   6330
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   585
         Width           =   2910
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
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "SKIPN"
         Top             =   585
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3030
         Picture         =   "frmICPostInventory.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   585
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Post Issue Note"
         Height          =   315
         Left            =   4125
         TabIndex        =   4
         Top             =   1740
         Width           =   2145
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   30
         TabIndex        =   2
         Top             =   1170
         Width           =   6255
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
            TabIndex        =   5
            Top             =   180
            Width           =   4185
         End
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Issue Note #  :"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   615
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Post Issue Note"
         Height          =   450
         Left            =   60
         TabIndex        =   3
         Top             =   210
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmICPostInvntory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ls_transcode As String
Dim ls_transcodePS As String
Dim ls_sql  As String
Dim PR_Dumy As New Recordset
Dim PR_Dumy1 As New Recordset
Dim ln_cnt
Dim PR_ICIssue As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object

Private Sub Command1_Click()
Dim ls_res
If Command1.Caption = "Post Issue Note" Then
    lblStatus = "Posting of Issue Note in Progress..."
    DoEvents
    ls_res = MsgBox("Post Voucher !!!", vbYesNo + vbInformation)
    If ls_res = vbYes Then
        If Trim(txtTransNo) = "" Then
            ls_res = MsgBox("Post ALL Issue Note Voucher !!!", vbYesNo + vbInformation)
            If ls_res = vbYes Then
                Call PostInventoryVoucher(Gs_compcode)
                Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
            End If
        Else
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
            txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
             PR_ICIssue.Open "select * from IC_IssueNoteMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         If PR_ICIssue.EOF Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtTransNo.SetFocus
         Else
             Call PostInventoryVoucher(Gs_compcode, txtTransNo)
             Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
         End If
        End If
        
    Else
        Call MsgBox("Voucher Not Posted !!!", vbInformation)
    End If
      
    lblStatus = ""

ElseIf Command1.Caption = "Post Issue Return Note" Then

    lblStatus = "Posting of Issue Return Note in Progress..."
    DoEvents
    ls_res = MsgBox("Post Voucher !!!", vbYesNo + vbInformation)
    If ls_res = vbYes Then
        If Trim(txtTransNo) = "" Then
            ls_res = MsgBox("Post ALL Issue return Note Voucher !!!", vbYesNo + vbInformation)
            If ls_res = vbYes Then
                Call PostIssueReturnVoucher(Gs_compcode)
                Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
            End If
        Else
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
            txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
             PR_ICIssue.Open "select * from IC_IssueReturnNoteMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         If PR_ICIssue.EOF Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtTransNo.SetFocus
         Else
             Call PostIssueReturnVoucher(Gs_compcode, txtTransNo)
             Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
         End If
        End If
        
    Else
        Call MsgBox("Voucher Not Posted !!!", vbInformation)
    End If
      
    lblStatus = ""

ElseIf Command1.Caption = "Post Adjustment Note" Then
    lblStatus = "Posting of Adjustment Note in Progress..."
    DoEvents
    ls_res = MsgBox("Post Voucher !!!", vbYesNo + vbInformation)
    If ls_res = vbYes Then
        If Trim(txtTransNo) = "" Then
            ls_res = MsgBox("Post ALL Adjustment Note Voucher !!!", vbYesNo + vbInformation)
            If ls_res = vbYes Then
                Call PostAdjustmentVoucher(Gs_compcode)
                Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
            End If
        Else
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
            txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
             PR_ICIssue.Open "select * from IC_InventoryAdjMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         If PR_ICIssue.EOF Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtTransNo.SetFocus
         Else
             Call PostAdjustmentVoucher(Gs_compcode, txtTransNo)
             Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
         End If
        End If
        
    Else
        Call MsgBox("Voucher Not Posted !!!", vbInformation)
    End If
    lblStatus = ""
ElseIf Command1.Caption = "Post Sale Consumption" Then
    lblStatus = "Posting of Sale Consumption in Progress..."
    DoEvents
    ls_res = MsgBox("Post Voucher !!!", vbYesNo + vbInformation)
    If ls_res = vbYes Then
        If Trim(txtTransNo) = "" Then
            ls_res = MsgBox("Post ALL Invoice Voucher !!!", vbYesNo + vbInformation)
            If ls_res = vbYes Then
                Call PostSaleConsumptionVoucher(Gs_compcode)
                Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
            End If
        Else
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
            txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
             PR_ICIssue.Open "select * from So_SaleInvoice where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus1 = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         If PR_ICIssue.EOF Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtTransNo.SetFocus
         Else
             Call PostSaleConsumptionVoucher(Gs_compcode, txtTransNo)
             Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
         End If
        End If
        
    Else
        Call MsgBox("Voucher Not Posted !!!", vbInformation)
    End If
    lblStatus = ""

End If
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtTransNo
    Set PO_DESC = Text1
    
    If Label1.Caption = "Post Issue Note" Then
        Gs_SQL = "Select TransCode, printtranscode,Transdate from IC_IssueNoteMaster "
        Gs_FindFld = "TransCode"
        Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "' and glstatus = 0"
        Gs_OrderBy = "Order by TransCode"
        MyLookupOLDB.Caption = "Issue Notes"
    ElseIf Label1.Caption = "Post Issue Return Note" Then
        Gs_SQL = "Select TransCode, printtranscode,Transdate from IC_IssueReturnNoteMaster "
        Gs_FindFld = "TransCode"
        Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "' and glstatus = 0"
        Gs_OrderBy = "Order by TransCode"
        MyLookupOLDB.Caption = "Issue Return Notes"
    ElseIf Label1.Caption = "Post Adjustment Note" Then
        Gs_SQL = "Select TransCode, printtranscode,Transdate from IC_InventoryAdjMaster "
        Gs_FindFld = "TransCode"
        Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "' and glstatus = 0"
        Gs_OrderBy = "Order by TransCode"
        MyLookupOLDB.Caption = "Adjustment Notes"
    ElseIf Label1.Caption = "Post Sale Consumption" Then
        Gs_SQL = "Select TransCode,Transdate from So_SaleInvoice "
        Gs_FindFld = "TransCode"
        Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "' and glstatus1 = 0"
        Gs_OrderBy = "Order by TransCode"
        MyLookupOLDB.Caption = "Sale Invoices"
    End If

    MyLookupOLDB.Show 1
    If txtTransNo <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn And Len(txtTransNo.Text) > 0 Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
         If Label1.Caption = "Post Issue Note" Then
            PR_ICIssue.Open "select * from IC_IssueNoteMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         ElseIf Label1.Caption = "Post Issue Return Note" Then
            PR_ICIssue.Open "select * from IC_IssueReturnNoteMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         ElseIf Label1.Caption = "Post Adjustment Note" Then
            PR_ICIssue.Open "select * from IC_InventoryAdjMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         ElseIf Label1.Caption = "Post Sale Consumption" Then
            PR_ICIssue.Open "select * from So_SaleInvoice where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus1 = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            
         End If
         If PR_ICIssue.EOF Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtTransNo.SetFocus
         Else
            If Label1.Caption = "Post Sale Consumption" Then
                Text1 = PR_ICIssue("Transdate")
            Else
                Text1 = PR_ICIssue("PrintTranscode")
            End If
         End If
     
  End If
End Sub
