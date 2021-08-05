VERSION 5.00
Begin VB.Form frmSOPaidAmtform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receive Amount"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSOPaidAmtform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkDiscAllowYN 
      Caption         =   "Discount Y/N"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   33
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox txtDiscAllowYN 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox txtBillCopy 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   8115
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   45
      TabIndex        =   22
      Top             =   -75
      Width           =   6000
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Invoice Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   60
         TabIndex        =   23
         Top             =   240
         Width           =   5850
      End
   End
   Begin VB.Frame Frame3 
      Height          =   660
      Left            =   45
      TabIndex        =   25
      Top             =   555
      Width           =   6000
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Bill Processing"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   50
         TabIndex        =   26
         Top             =   150
         Width           =   5895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6900
      Left            =   45
      TabIndex        =   7
      Top             =   1095
      Width           =   6000
      Begin VB.TextBox txtitemDiscounts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1020
         Width           =   2970
      End
      Begin VB.TextBox txtdiscby 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2790
         MaxLength       =   6
         TabIndex        =   24
         Top             =   2595
         Width           =   2955
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Disc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   5160
         TabIndex        =   21
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtdiscper 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2790
         TabIndex        =   19
         Top             =   1815
         Width           =   720
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   60
         TabIndex        =   6
         Top             =   6270
         Width           =   1530
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4275
         TabIndex        =   1
         Top             =   6270
         Width           =   1530
      End
      Begin VB.TextBox txtdisamount 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   3945
         TabIndex        =   2
         Top             =   1815
         Width           =   1200
      End
      Begin VB.TextBox txtnetamount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3060
         Width           =   2970
      End
      Begin VB.TextBox txtCreditCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   5
         Top             =   5385
         Width           =   1065
      End
      Begin VB.TextBox txtCreditDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   30
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   5880
         Width           =   5760
      End
      Begin VB.CommandButton Command5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3855
         Picture         =   "frmSOPaidAmtform.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   5370
         Width           =   360
      End
      Begin VB.CheckBox ChkCreditSale 
         Caption         =   "Credit &Sale"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4290
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1710
      End
      Begin VB.TextBox txtBalAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4605
         Width           =   2970
      End
      Begin VB.TextBox txtRecAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   2790
         TabIndex        =   0
         Top             =   3840
         Width           =   2970
      End
      Begin VB.TextBox txttotalamount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Items Discount:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   28
         Top             =   1155
         Width           =   2670
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3420
         TabIndex        =   20
         Top             =   1965
         Width           =   435
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount By :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   60
         TabIndex        =   18
         Top             =   2550
         Width           =   2670
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "&Discount Amount :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   75
         TabIndex        =   17
         Top             =   1920
         Width           =   2670
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Amount :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   16
         Top             =   3180
         Width           =   2235
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit ID :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   885
         TabIndex        =   15
         Top             =   5415
         Width           =   1830
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance  Amount :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   135
         TabIndex        =   14
         Top             =   4740
         Width           =   2580
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Receive Amount  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   13
         Top             =   3915
         Width           =   2580
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   9
         Top             =   330
         Width           =   2580
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Disc Per :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   32
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Bill Copy :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -120
      TabIndex        =   30
      Top             =   8160
      Width           =   1335
   End
End
Attribute VB_Name = "frmSOPaidAmtform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pr_dumy As New Recordset
Dim pr_dumy1 As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object
Dim ls_sql As String
Dim ls_ClientCreditCode As String

Dim ls_res
Private Sub ChkCreditSale_Click()
If ChkCreditSale.Value = 1 Then
    txtCreditCode.Enabled = True
    txtCreditDesc.Enabled = True
    Command5.Enabled = True
    
    txtCreditCode.SetFocus
Else
    txtCreditCode = ""
    txtCreditDesc = ""
    txtCreditCode.Enabled = False
    txtCreditDesc.Enabled = False
    Command5.Enabled = False

End If
End Sub


Private Sub Command1_Click()

Command1.Enabled = False

On Error GoTo LocalErr
Dim ln_dicamount As Double
Dim ls_clientcode As String
Dim ls_transcode As String
If Val(txtrecamount) < Val(txtNetAmount) Then
   Call MsgBox("Amount not Valid !!!", vbCritical)
   txtrecamount.SetFocus
    Exit Sub
Else
    If ChkCreditSale.Value = 1 And Trim(txtCreditCode) = "" Then
        Call MsgBox("Must Enter/Select Credit Code !!!", vbCritical)
        txtCreditCode.SetFocus
        Exit Sub
 Else
 
 DoEvents
'gc_dbcon.BeginTrans

With frmSO_Posform
    .txtstatus = "OK"
    .txttransno.Text = maxtranscode

    frmSO_Posform.dtptransdate = Gd_SysDate

    frmSO_Posform.Refresh


    Label10.Caption = "Saving Records..."
    DoEvents

If ChkCreditSale.Value = 1 Then
    ls_clientcode = txtCreditCode
    txtrecamount = txtNetAmount
Else
    ls_clientcode = "000019"
End If
'
DoEvents
frmSO_Posform.Refresh

ls_sql = "Insert into SO_TransMaster(Compcode, TransCode, TransDate, AccountCode, Remarks, TotalAmount, DiscPer, DiscAmount, NetAmount, RecAmount, BalAmount,DiscBy,SaleStatus,usercode,compname)"
ls_sql = ls_sql & " Values ('" & Gs_compcode & "' , '" & frmSO_Posform.txttransno & "', '" & Format(.dtptransdate, "YYYY/MM/DD HH:MM:SS") & "', '" & ls_clientcode & "' , 'Sale' , " & Val(txttotalamount) & ", " & Val(txtdiscper) & ", " & Val(txtdisamount) & ", " & Val(txtNetAmount) & ", " & Val(txtrecamount) & ", " & Val(txtBalAmount) & ",'" & txtdiscby.Text & "'," & ChkCreditSale.Value & "," & Gn_UserCode & ",'" & Gs_ComputerName & "')"
gc_dbcon.Execute ls_sql

DoEvents
Dim mRemarks As String
If POption = "R1" Then
   mRemarks = "Ramdan Package 1"
ElseIf POption = "R2" Then
   mRemarks = "Ramdan Package 2"
Else
  mRemarks = "SALE"
End If
 With frmSO_Posform.GrdGRN
       For ln_cnt = 1 To .Rows - 1
       If .TextMatrix(ln_cnt, 1) <> "" And Val(.TextMatrix(ln_cnt, 3)) > 0 And Val(.TextMatrix(ln_cnt, 5)) > 0 Then
        
        ln_dicamount = ((Val(.TextMatrix(ln_cnt, 5)) / Val(txttotalamount)) * Val(txtdisamount))
        
        
        ls_sql = "INSERT into SO_Trans(Compcode, TransCode,customcode, ItemCode, Quantity,Itemrate,Amount,discamount,AvgRate,packqty,packdisc,Remarks,LPRate,CatCode,LNPRate)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & frmSO_Posform.txttransno & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 10)) & "'," & (Val(.TextMatrix(ln_cnt, 3))) & "," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & "," & Val(.TextMatrix(ln_cnt, 9)) + ln_dicamount + Val(.TextMatrix(ln_cnt, 19)) & "," & Val(.TextMatrix(ln_cnt, 14)) & "," & Val(.TextMatrix(ln_cnt, 18)) & "," & Val(.TextMatrix(ln_cnt, 19)) & ",'" & mRemarks & "'," & Val(.TextMatrix(ln_cnt, 24)) & ",'" & .TextMatrix(ln_cnt, 25) & "'," & Val(.TextMatrix(ln_cnt, 26)) & ")"
        gc_dbcon.Execute ls_sql
        
        DoEvents
        
      End If
      Next
  End With
  
' gc_dbcon.CommitTrans
 
' With frmSO_Posform
'       'MSComm1.Output = txtitemdesc & Chr$(13)
'        ls_Dispname = Space(5) & "THANK YOU"
'        ln_strlen = 20 - Len(ls_Dispname)
'        ls_Dispname = ls_Dispname + Space(ln_strlen)
'        If .MSComm1.PortOpen Then .MSComm1.PortOpen = False
'        On Error Resume Next
'        .MSComm1.PortOpen = True
'        .MSComm1.Output = Space(40) + Chr$(13)
'        .MSComm1.Output = Space(2) & "FOR YOUR VISITING " & ls_Dispname & Chr$(10)       ' Ensure that
'        .MSComm1.PortOpen = False
'
'End With

    Label10.Caption = ""
     frmSO_Posform.txtBillCopy = txtBillCopy.Text
    Unload Me

    End With
    Exit Sub
End If
End If
LocalErr:
Call MsgBox(Err.Description)
End Sub
Private Function maxtranscode() As String
 If pr_dumy.State = 1 Then pr_dumy.Close
pr_dumy.Open "select max(transcode) as transcode from SO_TransMaster where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function
Private Sub Command2_Click()
frmSO_Posform.txtstatus = "Cancel"
Unload Me
End Sub

Private Sub File_Click()

End Sub

Private Sub Command3_Click()
frmpasswordform.txtopt = 1
frmpasswordform.Show 1
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
frmSO_Posform.txtstatus = "Cancel"
Unload Me
End If
End Sub

Private Sub Form_Load()
Command1.Enabled = True
txtdiscper.Enabled = True
'If frmSO_Posform.txtMedicin.Text <> "U GARMENTS" Then
' If Time() >= "9:00:00 PM" Then
'  txtdiscper = 10
  
' ElseIf Time() < "9:00:00 PM" Then
'  txtdiscper = 0
' End If
'End If
'txtdiscper.Enabled = False
End Sub

Private Sub Label4_Click()
txtdiscper.SetFocus
End Sub

Private Sub txtCreditCode_LostFocus()
If txtCreditCode <> "" Then
txtrecamount = txtNetAmount
txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
End If
End Sub

Private Sub txtdisamount_Change()

txtNetAmount = Val(txttotalamount) - (Val(txtdisamount) + Val(txtitemDiscounts))
If Val(txtNetAmount) < Val(txtrecamount) Then
    txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Else
    txtBalAmount = Val(txtNetAmount) - Val(txtrecamount)
End If

End Sub

Private Sub txtdisamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Val(txtdisamount) = 0 Then
        txtrecamount.SetFocus
    Else
        txtrecamount.SetFocus
    End If
End If
End Sub

Private Sub txtdisamount_LostFocus()
'If txtdisamount <> "" And frmSO_Posform.ChkEmpBill.Value = 1 Then
'  If Val(txtdisamount) > Val(frmSO_Posform.txtempDisc) Then
'    ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
'    If ls_res = vbNo Then
'     txtdisamount = ""
'     txtdisamount.SetFocus
'    End If
'  Else
'     txtdisamount = ""
'  End If
'
'ElseIf txtdisamount <> "" Then
'    If Val(txtdisamount) > Val(frmSO_Posform.txtdiscamount) Then
'    ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
'    If ls_res = vbNo Then
'     txtdisamount = ""
'     txtdisamount.SetFocus
'    End If
'  End If
'Else
'  txtdisamount = ""
'End If
'

txtNetAmount = Val(txttotalamount) - (Val(txtdisamount) + Val(txtitemDiscounts))

'txtnetamount = Val(txttotalamount) - (Val(txtdisamount) + Val(txtitemDiscounts))

If Val(txtNetAmount) < Val(txtrecamount) Then
    txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Else
    txtBalAmount = Val(txtNetAmount) - Val(txtrecamount)
End If
End Sub


Private Sub txtdiscper_Change()
On Error GoTo LocalErr
If Trim(txtdiscper) <> "" Then

'txtdisamount = Round(Val(txtNetAmount) * Val(txtdiscper) / 100, 0)


txtdisamount = Round(Val(txttotalamount) * Val(txtdiscper) / 100, 0)
Else
txtdisamount = ""
End If

If txtdisamount <> "" And frmSO_Posform.ChkEmpBill.Value = 1 Then
 ' If Val(txtdisamount) > Val(frmSO_Posform.txtempDisc) Then
 '   ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
 '   If ls_res = vbNo Then
 '    txtdiscper = 0
 '    txtdiscper.SetFocus
 '   End If
 ' End If

ElseIf txtdisamount <> "" Then
  '  If Val(txtdisamount) > Val(frmSO_Posform.txtdiscamount) Then
  '  ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
  '  If ls_res = vbNo Then
  '   txtdiscper = 0
  '   txtdiscper.SetFocus
  '  End If
  'End If
Else
  txtdisamount = ""
End If
txtNetAmount = Val(txttotalamount) - (Val(txtdisamount) + Val(txtitemDiscounts))

If Val(txtNetAmount) < Val(txtrecamount) Then
    txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Else
    txtBalAmount = Val(txtNetAmount) - Val(txtrecamount)
End If
       If gn_comportset > 0 Then
          With frmSO_Posform
              On Error GoTo 0
           
               If .MSComm1.PortOpen Then MSComm1.PortOpen = False
               .MSComm1.CommPort = gn_comportset
               .MSComm1.Settings = "9600,N,8,1"
               .MSComm1.InputLen = 0
               .MSComm1.PortOpen = True
                ls_Dispname = "Total:"
                .MSComm1.Output = Space(40) + Chr$(13)
                .MSComm1.Output = ls_Dispname & str(Val(txtNetAmount)) & Chr$(13) & Chr$(10)   ' Ensure that
               .MSComm1.PortOpen = False
            End With
            End If

Exit Sub
LocalErr:
Call MsgBox(Err.Description)
End Sub

Private Sub txtdiscper_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    If Trim(txtdiscper) = "" Then
        txtdisamount.Enabled = True
        txtdisamount.SetFocus
    Else
        txtrecamount.SetFocus
    End If
End If
End Sub

Private Sub txtdiscper_LostFocus()
On Error GoTo LocalErr
If Trim(txtdiscper) <> "" Then
txtdisamount = Round(Val(txtNetAmount) * Val(txtdiscper) / 100, 0)

'txtdisamount = Round(Val(txttotalamount) * Val(txtdiscper) / 100, 0)


End If

If txtdisamount <> "" And frmSO_Posform.ChkEmpBill.Value = 1 Then
 ' If Val(txtdisamount) > Val(frmSO_Posform.txtempDisc) Then
 '   ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
 '   If ls_res = vbNo Then
 '    txtdiscper = 0
 '    txtdiscper.SetFocus
 '   End If
 ' End If

ElseIf txtdisamount <> "" Then
  '  If Val(txtdisamount) > Val(frmSO_Posform.txtdiscamount) Then
  '  ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
  '  If ls_res = vbNo Then
  '   txtdiscper = 0
  '   txtdiscper.SetFocus
  '  End If
  'End If
Else
  txtdisamount = ""
End If
txtNetAmount = Val(txttotalamount) - (Val(txtdisamount) + Val(txtitemDiscounts))

If Val(txtNetAmount) < Val(txtrecamount) Then
    txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Else
    txtBalAmount = Val(txtNetAmount) - Val(txtrecamount)
End If
       If gn_comportset > 0 Then
          With frmSO_Posform
              On Error GoTo 0
           
               If .MSComm1.PortOpen Then MSComm1.PortOpen = False
               .MSComm1.CommPort = gn_comportset
               .MSComm1.Settings = "9600,N,8,1"
               .MSComm1.InputLen = 0
               .MSComm1.PortOpen = True
                ls_Dispname = "Total:"
                .MSComm1.Output = Space(40) + Chr$(13)
                .MSComm1.Output = ls_Dispname & str(Val(txtNetAmount)) & Chr$(13) & Chr$(10)   ' Ensure that
               .MSComm1.PortOpen = False
            End With
            End If

Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub

Private Sub txtRecAmount_Change()


If txtrecamount <> "" Then
    If Not IsNumeric(txtrecamount) Then
        Call MsgBox("Numeric entery only !!!", vbCritical)
        txtrecamount = ""
        txtrecamount.SetFocus
        Exit Sub
    End If
    
    
    
    If Val(txtNetAmount) < Val(txtrecamount) Then
        txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
    Else
        txtBalAmount = Val(txtNetAmount) - Val(txtrecamount)
    End If
    
    'txtBalAmount = Val(txtRecAmount) - Val(txtnetamount)
  
 
    
ElseIf txtrecamount = "" Then
    txtrecamount = ""
End If

   
End Sub
Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCreditCode
    Set PO_DESC = txtCreditDesc
    Gs_ExtraPara = ""
    Gs_SQL = "Select Clientcode, Description from IC_clients "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Credit Clients"
    MyLookupOLDB.Show 1
    
    If txtCreditCode <> "" Then Call txtCreditCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtCreditCode_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LocalErr
ChkDiscAllowYN.Enabled = True
If Trim(txtCreditCode) <> "" And KeyCode = vbKeyReturn Then
        
        If pr_dumy.State = 1 Then pr_dumy.Close
        txtCreditCode = DoPad(txtCreditCode, 6)
        pr_dumy.Open "Select * from IC_clients where Compcode  = '" & Gs_compcode & "' and clientcode= '" & txtCreditCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Client Code not found !!!", vbCritical)
            txtCreditCode = ""
            txtCreditCode = ""
            txtCreditCode.SetFocus
        Else
            txtCreditDesc = pr_dumy("Description")
            
            txtrecamount = txtNetAmount
            txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
            txtBillCopy = pr_dumy("BillCopy")
            txtDiscAllowYN = pr_dumy("DiscAllowYN")
            ChkDiscAllowYN.Value = pr_dumy("DiscAllowYN")
            ChkDiscAllowYN.Enabled = False
            Command1.SetFocus
        End If
        
If Val(txtCreditCode) = 47 Then

'If Val(pr_dumy("DiscAllowYN") >= 1 Then
   'This Proceding for Grocer App Marketing
  
frmSO_Posform.txttotalamount = ""
frmSO_Posform.txttotalqty = ""
frmSO_Posform.txtdiscamount = ""
frmSO_Posform.txtempDisc = ""
frmSO_Posform.txtpackdisc = ""
frmSO_Posform.txtNetAmount = ""
'txtMedicin = ""
'txtOfferAmnt = ""
 ln_cnt = 0
  
  With frmSO_Posform.GrdGRN
     For ln_cnt = 1 To .Rows - 1
        .TextMatrix(ln_cnt, 9) = 0
         frmSO_Posform.txttotalamount = Format(Val(frmSO_Posform.txttotalamount) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
         frmSO_Posform.txttotalqty = Format(Val(frmSO_Posform.txttotalqty) + Val(.TextMatrix(ln_cnt, 3)), "######0.00")
         frmSO_Posform.txtdiscamount = Format(Val(frmSO_Posform.txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), "######0.00")
         frmSO_Posform.txtempDisc = Format(Val(frmSO_Posform.txtempDisc) + Val(.TextMatrix(ln_cnt, 11)), "######0.00")
         frmSO_Posform.txtpackdisc = Format(Val(frmSO_Posform.txtpackdisc) + Val(.TextMatrix(ln_cnt, 19)), "######0.00")
          frmSO_Posform.txtNetAmount = Format(Val(frmSO_Posform.txttotalamount) - Val(frmSO_Posform.txtdiscamount), "######0.00")
     Next
  End With

txtitemDiscounts = 0
txtdiscper = 0
txtdisamount = frmSO_Posform.txtdiscamount
txttotalamount = frmSO_Posform.txtNetAmount
txtNetAmount = Val(txttotalamount) - (Val(txtdisamount) + Val(txtitemDiscounts))
txtrecamount = txtNetAmount
txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Command1.SetFocus

ElseIf Val(txtCreditCode) = 46 Then

   'This Proceding for Police traning College

frmSO_Posform.txttotalamount = ""
frmSO_Posform.txttotalqty = ""
frmSO_Posform.txtdiscamount = ""
frmSO_Posform.txtempDisc = ""
frmSO_Posform.txtpackdisc = ""
frmSO_Posform.txtNetAmount = ""
'txtMedicin = ""
'txtOfferAmnt = ""
 ln_cnt = 0
txtitemDiscounts = ""
  
  With frmSO_Posform.GrdGRN
     For ln_cnt = 1 To .Rows - 1
        If Val(.TextMatrix(ln_cnt, 25)) <= 10 Or Val(.TextMatrix(ln_cnt, 25)) = 49 Or Val(.TextMatrix(ln_cnt, 25)) = 50 Or Val(.TextMatrix(ln_cnt, 25)) = 51 Then
           .TextMatrix(ln_cnt, 9) = ((Val(.TextMatrix(ln_cnt, 5)) / 100) * 25)
        Else
         If Val(.TextMatrix(ln_cnt, 25)) <> 21 Then
           .TextMatrix(ln_cnt, 4) = Round(Val((.TextMatrix(ln_cnt, 26) + (Val(.TextMatrix(ln_cnt, 26)) / 100) * 2)), 2)
           .TextMatrix(ln_cnt, 4) = Round(Val(.TextMatrix(ln_cnt, 4)), 2)
           .TextMatrix(ln_cnt, 5) = Round(Val(.TextMatrix(ln_cnt, 3)) * Val(.TextMatrix(ln_cnt, 4)), 2)
           .TextMatrix(ln_cnt, 9) = 0
         End If
           .TextMatrix(ln_cnt, 9) = 0
        End If
         frmSO_Posform.txttotalamount = Format(Val(frmSO_Posform.txttotalamount) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
         frmSO_Posform.txttotalqty = Format(Val(frmSO_Posform.txttotalqty) + Val(.TextMatrix(ln_cnt, 3)), "######0.00")
         frmSO_Posform.txtdiscamount = Format(Val(frmSO_Posform.txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), "######0.00")
         frmSO_Posform.txtempDisc = Format(Val(frmSO_Posform.txtempDisc) + Val(.TextMatrix(ln_cnt, 11)), "######0.00")
         frmSO_Posform.txtpackdisc = Format(Val(frmSO_Posform.txtpackdisc) + Val(.TextMatrix(ln_cnt, 19)), "######0.00")
         frmSO_Posform.txtNetAmount = Format(Val(frmSO_Posform.txttotalamount) - Val(frmSO_Posform.txtdiscamount), "######0.00")
     Next
  End With

txtitemDiscounts = frmSO_Posform.txtdiscamount
txttotalamount = frmSO_Posform.txttotalamount
txtNetAmount = Val(txttotalamount) - (Val(txtdisamount) + Val(txtitemDiscounts))
txtrecamount = txtNetAmount
txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Command1.SetFocus

End If

ElseIf Trim(txtCreditCode) = "" And KeyCode = vbKeyReturn Then
        txtCreditCode = ""
        txtCreditDesc = ""
        Call Command5_Click
End If

Exit Sub
LocalErr:
Call MsgBox(Err.Description, vbCritical)

End Sub

Private Sub txtrecamount_GotFocus()
On Error GoTo LocalErr
If Trim(txtdiscper) <> "" Then
txtdisamount = Round(Val(txtNetAmount) * Val(txtdiscper) / 100, 0)

'txtdisamount = Round(Val(txttotalamount) * Val(txtdiscper) / 100, 0)




End If

If txtdisamount <> "" And frmSO_Posform.ChkEmpBill.Value = 1 Then
 ' If Val(txtdisamount) > Val(frmSO_Posform.txtempDisc) Then
 '   ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
 '   If ls_res = vbNo Then
 '    txtdiscper = 0
 '    txtdiscper.SetFocus
 '   End If
 ' End If

ElseIf txtdisamount <> "" Then
  '  If Val(txtdisamount) > Val(frmSO_Posform.txtdiscamount) Then
  '  ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
  '  If ls_res = vbNo Then
  '   txtdiscper = 0
  '   txtdiscper.SetFocus
  '  End If
  'End If
Else
  txtdisamount = ""
End If
txtNetAmount = Val(txttotalamount) - (Val(txtdisamount) + Val(txtitemDiscounts))

If Val(txtNetAmount) < Val(txtrecamount) Then
    txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Else
    txtBalAmount = Val(txtNetAmount) - Val(txtrecamount)
End If
       If gn_comportset > 0 Then
          With frmSO_Posform
              On Error GoTo 0
           
               If .MSComm1.PortOpen Then MSComm1.PortOpen = False
               .MSComm1.CommPort = gn_comportset
               .MSComm1.Settings = "9600,N,8,1"
               .MSComm1.InputLen = 0
               .MSComm1.PortOpen = True
                ls_Dispname = "Total:"
                .MSComm1.Output = Space(40) + Chr$(13)
                .MSComm1.Output = ls_Dispname & str(Val(txtNetAmount)) & Chr$(13) & Chr$(10)   ' Ensure that
               .MSComm1.PortOpen = False
            End With
            End If
txtdisamount.Enabled = False
txtdiscper.Enabled = False
Exit Sub
LocalErr:
Call MsgBox(Err.Description)


End Sub

Private Sub txtRecAmount_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LocalErr
If KeyCode = vbKeyReturn Then
Command1.Enabled = True
If Val(txtNetAmount) < Val(txtrecamount) Then
    txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Else
    txtBalAmount = Val(txtNetAmount) - Val(txtrecamount)
End If



 With frmSO_Posform
       'MSComm1.Output = txtitemdesc & Chr$(13)
       ls_Dispname = "Received:" + Trim(str(Val(txtrecamount.Text)) & "")
       ln_strlen = 20 - Len(ls_Dispname)
       ls_Dispname = ls_Dispname + Space(ln_strlen)
       
       On Error Resume Next
       If .MSComm1.PortOpen Then .MSComm1.PortOpen = False

      .MSComm1.PortOpen = True
      '.MSComm1.Output = Space(40) + Chr$(13)
      .MSComm1.Output = ls_Dispname & "Balance:" + Trim(str(Val(Round(txtBalAmount.Text, 0))) & "") & Chr$(13) & Chr$(10) ' Ensure that
      .MSComm1.PortOpen = False
   
End With

Command1.SetFocus
End If

Exit Sub
LocalErr:
Call MsgBox(Err.Description, vbCritical)

End Sub

Private Sub txtrecamount_LostFocus()
If txtrecamount <> "" Then
    If Val(txtNetAmount) < Val(txtrecamount) Then
    txtBalAmount = Val(txtrecamount) - Val(txtNetAmount)
Else
    txtBalAmount = Val(txtNetAmount) - Val(txtrecamount)
End If
End If
End Sub

