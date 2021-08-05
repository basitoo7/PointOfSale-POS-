VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSOPostSaleFTDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post Sale Voucher"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "frmSOPostSaleFTDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   4920
      Begin VB.CheckBox chkrepost 
         Caption         =   "Re-Post"
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   1605
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Post Sale Voucher"
         Height          =   315
         Left            =   2730
         TabIndex        =   2
         Top             =   1575
         Width           =   2130
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   45
         TabIndex        =   1
         Top             =   1005
         Width           =   4830
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
            TabIndex        =   3
            Top             =   180
            Width           =   4185
         End
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1575
         TabIndex        =   4
         Top             =   165
         Width           =   1440
         _ExtentX        =   2540
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
         CustomFormat    =   "dd-MM-yyyy HH:mm:ss"
         Format          =   64684033
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker DTPto 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   570
         Width           =   1440
         _ExtentX        =   2540
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
         CustomFormat    =   "dd-MM-yyyy HH:mm:ss"
         Format          =   64684033
         CurrentDate     =   37293
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To Date :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   810
         TabIndex        =   7
         Top             =   585
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   690
         TabIndex        =   5
         Top             =   180
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmSOPostSaleFTDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Me.Caption = "Post Sale GL Voucher" Then

    If chkrepost.Value = 1 Then
       ls_sql = "Delete from gl_ref where vchrtype = 'SVS' and value_date >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Value_date <='" & Format(dtpto, "YYYY/MM/DD") & "' "
       gc_dbcon.Execute ls_sql
       ls_sql = "Delete from gl_trans where vchrtype = 'SVS' and value_date >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Value_date <='" & Format(dtpto, "YYYY/MM/DD") & "' "
       gc_dbcon.Execute ls_sql
       ls_sql = " update  SO_TransMaster set glstatus =0"
       ls_sql = ls_sql & " Where convert(varchar,transdate,111) >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' and convert(varchar,transdate,111) <= '" & Format(dtpto, "YYYY/MM/DD") & "' "
       gc_dbcon.Execute ls_sql
    
    
    End If
   
    If PostSaleVoucher(Gs_compcode, dtpfrom, dtpto) Then
     Call MsgBox("Sale Voucher Successfully Posted !!!", vbInformation)
     Else
     Call MsgBox("Voucher not Posted !!!", vbCritical)
    End If
   
   
    
 ElseIf Me.Caption = "Post Cost Of Sale" Then
    
    UpdateCostofNonAvgRate
    If chkrepost.Value = 1 Then
       ls_sql = "Delete from gl_ref where vchrtype = 'CSV' and value_date >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Value_date <='" & Format(dtpto, "YYYY/MM/DD") & "' "
       gc_dbcon.Execute ls_sql
       ls_sql = "Delete from gl_trans where vchrtype = 'CSV' and value_date >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Value_date <='" & Format(dtpto, "YYYY/MM/DD") & "' "
       gc_dbcon.Execute ls_sql
    End If
    
    If PostCostSaleVoucher(Gs_compcode, dtpfrom, dtpto) Then
        Call MsgBox("Cost Voucher Successfully Posted !!!", vbInformation)
    End If
'
ElseIf Me.Caption = "Post Credit Sale GL Voucher" Then
    If PostSaleCreditVoucherCustomer(Gs_compcode, dtpfrom, dtpto) Then
     Call MsgBox("Sale Voucher Successfully Posted !!!", vbInformation)
     Else
     Call MsgBox("Voucher not Posted !!!", vbCritical)
    End If
ElseIf Me.Caption = "Post Sale Return GL Voucher" Then

   If chkrepost.Value = 1 Then
       ls_sql = "Delete from gl_ref where vchrtype = 'SRS' and value_date >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Value_date <='" & Format(dtpto, "YYYY/MM/DD") & "' "
       gc_dbcon.Execute ls_sql
       ls_sql = "Delete from gl_trans where vchrtype = 'SRS' and value_date >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Value_date <='" & Format(dtpto, "YYYY/MM/DD") & "' "
       gc_dbcon.Execute ls_sql
       ls_sql = " update  SO_TransReturnMaster set glstatus =0"
       ls_sql = ls_sql & " Where convert(varchar,transdate,111) >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' and convert(varchar,transdate,111) <= '" & Format(dtpto, "YYYY/MM/DD") & "' "
       gc_dbcon.Execute ls_sql
       
      
    End If
  


    If PostSaleReturnVchr(Gs_compcode, dtpfrom, dtpto) Then
     Call MsgBox("Sale Return Voucher Successfully Posted !!!", vbInformation)
     Else
     Call MsgBox("Voucher not Posted !!!", vbCritical)
    End If
'
'   If PostCostReturnVoucher(Gs_compcode, dtpfrom, DTPto) Then
'     Call MsgBox("Cost of Sale Return Voucher Successfully Posted !!!", vbInformation)
'     Else
'     Call MsgBox("Voucher not Posted !!!", vbCritical)
'    End If
End If

End Sub

Private Sub Form_Load()
  dtpfrom = Date
  dtpto = Date
End Sub
