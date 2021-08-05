VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmicreport13 
   Caption         =   "Update Purchase Record"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports13.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Post Voucher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      MaskColor       =   &H00000000&
      TabIndex        =   18
      Top             =   2595
      Width           =   1545
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H00000080&
      Height          =   2580
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7965
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   6465
         TabIndex        =   19
         Top             =   600
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52822017
         CurrentDate     =   41576
      End
      Begin VB.TextBox txtoldclientcode 
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   16
         Top             =   975
         Width           =   690
      End
      Begin VB.TextBox txtClientCode 
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   14
         Top             =   1335
         Width           =   690
      End
      Begin VB.TextBox txtclientdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2745
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1335
         Width           =   5100
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2400
         Picture         =   "frmicreports13.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1335
         Width           =   315
      End
      Begin VB.TextBox txtremarks 
         Height          =   315
         Left            =   1680
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1770
         Width           =   6180
      End
      Begin VB.ComboBox txtstatus 
         Height          =   330
         ItemData        =   "frmicreports13.frx":047C
         Left            =   1680
         List            =   "frmicreports13.frx":048C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   2355
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   3045
         Picture         =   "frmicreports13.frx":04D1
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   615
         Width           =   315
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   3375
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   615
         Width           =   3060
      End
      Begin VB.TextBox txtTransNo 
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label lbloldclient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2430
         TabIndex        =   20
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Old Client Code :"
         Height          =   210
         Left            =   450
         TabIndex        =   17
         Top             =   1005
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Supplier Code :"
         Height          =   210
         Left            =   555
         TabIndex        =   15
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Remarks :"
         Height          =   210
         Left            =   825
         TabIndex        =   11
         Top             =   1785
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Change Status of :"
         Height          =   210
         Left            =   210
         TabIndex        =   10
         Top             =   225
         Width           =   1350
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Trans No :"
         Height          =   210
         Left            =   795
         TabIndex        =   9
         Top             =   645
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5835
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   2625
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Exit"
      Height          =   330
      Left            =   6930
      TabIndex        =   4
      Top             =   2640
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2985
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Description"
            TextSave        =   "Description"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   105833
            MinWidth        =   105833
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmicreport13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Dumy As New Recordset
Dim ls_sql As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
If txtstatus = "" Then
Call MsgBox("Please Select Status For !!!", vbCritical)
txtstatus.SetFocus
ElseIf txtTransNo = "" Then
Call MsgBox("Please Select/Enter Document No !!!", vbCritical)
txtTransNo.SetFocus
ElseIf txtremarks = "" Then
Call MsgBox("Please Enter Remarks !!!", vbCritical)
txtremarks.SetFocus
Else


       ls_sql = "Insert into PO_RejRemarks (Compcode,Transcode,Transdate,Remarks) Values ( '" & Gs_compcode & "' ,'" & txtTransNo & "','" & Format(Date, "YYYY/MM/DD") & "' ,'" & txtremarks & "')"
       gc_dbcon.Execute ls_sql
         
       If txtstatus.ListIndex = 0 Then
            ls_sql = "update PO_POGRN set glstatus =0"
            ls_sql = ls_sql & " where Compcode ='" & Gs_compcode & "'  and transcode = '" & txtTransNo & "'"
            gc_dbcon.Execute ls_sql
            
            
            If txtClientCode <> "" Then
            
            ls_sql = "update PO_POGRN set accountcode = '" & txtClientCode & "'"
            ls_sql = ls_sql & " where Compcode ='" & Gs_compcode & "'  and transcode = '" & txtTransNo & "' "
            gc_dbcon.Execute ls_sql
            End If
                 
            DeleteVoucher
            
       
       ElseIf txtstatus.ListIndex = 2 Then
            ls_sql = "update PO_PayableMaster set glstatus =0"
            ls_sql = ls_sql & " where Compcode ='" & Gs_compcode & "'  and transcode = '" & txtTransNo & "'"
            gc_dbcon.Execute ls_sql
            DeleteBankVoucher
       ElseIf txtstatus.ListIndex = 3 Then
            ls_sql = "update PO_PayableCashMaster set glstatus =0"
            ls_sql = ls_sql & " where Compcode ='" & Gs_compcode & "'  and transcode = '" & txtTransNo & "'"
            gc_dbcon.Execute ls_sql
            DeleteCashVoucher
       Else
              ls_sql = "update PO_POGRNReturn set glstatus =0,status = 0"
              ls_sql = ls_sql & " where Compcode ='" & Gs_compcode & "'  and transcode = '" & txtTransNo & "'"
              gc_dbcon.Execute ls_sql
              
              
              If txtClientCode <> "" Then
              
              ls_sql = "update PO_POGRNReturn set accountcode = '" & txtClientCode & "'"
              ls_sql = ls_sql & " where Compcode ='" & Gs_compcode & "'  and transcode = '" & txtTransNo & "' "
              gc_dbcon.Execute ls_sql
              End If
                   
              DeleteGRRNVoucher
            
       End If
        
        
        Call MsgBox("Successfully Updated", vbInformation)
End If
'txtremarks = ""
'txtTransNo = ""
'txtdesc = ""


End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtClientCode
    Set PO_DESC = txtclientdesc
    
    Gs_SQL = "SELECT IC_Supplier.SupplierCode, IC_Supplier.Description, Cities.CityName, Tehseels.TehseelName, IC_Supplier.PhoneOffice, IC_Supplier.Address,"
    Gs_SQL = Gs_SQL & " IC_Supplier.CreditLimit FROM IC_Supplier LEFT OUTER JOIN   Cities ON IC_Supplier.CityCode = Cities.CityCode LEFT OUTER JOIN"
    Gs_SQL = Gs_SQL & " Tehseels ON IC_Supplier.CityCode = Tehseels.CityCode AND IC_Supplier.TehseelCode = Tehseels.TehseelCode"
                      
    
    
    
    Gs_FindFld = "IC_Supplier.Description"
    Gs_OrderBy = "Order by IC_Supplier.Description"
    Gs_OtherPara = " where IC_Supplier.Compcode ='" & Gs_compcode & "' "
    
    
'    gn_lcolw = 3500
'    gn_lcolw1 = 2000
'    gn_lcolw2 = 2000
'    gn_lcolw3 = 2000
'    gn_lcolw4 = 1500
'    gn_lcolw5 = 2000
'
    
    MyLookupOLDB.Caption = "Clients"
    MyLookupOLDB.Show 1
    
    If txtClientCode <> "" Then Call txtClientCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command2_Click()
If txtTransNo <> "" Then

If PostPurchaseVoucher(Gs_compcode, txtvaluedate, txtvaluedate, txtTransNo) Then
  Call MsgBox(" Purchase Voucher Successfully Posted !!!", vbInformation)
End If
Else
Call MsgBox("Please Select/Enter Purchase Invoice", vbCritical)
txtTransNo.SetFocus
End If

End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtTransNo
    Set PO_DESC = txtdesc
    
    If txtstatus.ListIndex = 0 Then
    Gs_SQL = "SELECT PO_POGRN.TransCode, IC_Supplier.Description FROM  PO_POGRN LEFT OUTER JOIN "
    Gs_SQL = Gs_SQL & " IC_Supplier ON PO_POGRN.Compcode = IC_Supplier.Compcode AND PO_POGRN.AccountCode = IC_Supplier.SupplierCode"
    Gs_FindFld = " IC_Supplier.Description"
    'Gs_OtherPara = " Where PO_POGRN.glstatus = 1 "
    Gs_OrderBy = "Order by PO_POGRN.TransCode Desc"
    MyLookupOLDB.Caption = "GRN Notes"
    ElseIf txtstatus.ListIndex = 1 Then
    Gs_SQL = "SELECT PO_POGRNReturn.TransCode, IC_Supplier.Description FROM  PO_POGRNReturn LEFT OUTER JOIN "
    Gs_SQL = Gs_SQL & " IC_Supplier ON PO_POGRNReturn.Compcode = IC_Supplier.Compcode AND PO_POGRNReturn.AccountCode = IC_Supplier.SupplierCode"
    Gs_FindFld = " IC_Supplier.Description"
   ' Gs_OtherPara = " Where PO_POGRNReturn.glstatus = 1 "
    Gs_OrderBy = "Order by PO_POGRNReturn.TransCode Desc"
    MyLookupOLDB.Caption = "GRRN Notes"
    
    ElseIf txtstatus.ListIndex = 2 Then
    
    Gs_SQL = "Select TransCode, Transdate from PO_PayableMaster"
    Gs_FindFld = "TransCode"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by TransCode"
    MyLookupOLDB.Caption = "Bank Payment #"
    
    ElseIf txtstatus.ListIndex = 3 Then
    
    Gs_SQL = "Select TransCode, Transdate from PO_PayableCashMaster"
    Gs_FindFld = "TransCode"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by TransCode"
    MyLookupOLDB.Caption = "Cash Payment #"
   
    
    End If
    
    MyLookupOLDB.Show 1
    If txtTransNo <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)


End Sub

Private Sub DeleteVoucher()
ls_sql = "Delete from gl_ref where compcode = '" & Gs_compcode & "' and vchrtype = 'PVS' and Voucher_no = '" & txtTransNo & "'"
gc_dbcon.Execute ls_sql
ls_sql = "Delete from gl_trans where compcode = '" & Gs_compcode & "' and vchrtype = 'PVS' and Voucher_no = '" & txtTransNo & "'"
gc_dbcon.Execute ls_sql
End Sub

Private Sub DeleteGRRNVoucher()
ls_sql = "Delete from gl_ref where compcode = '" & Gs_compcode & "' and vchrtype = 'PRV' and Voucher_no = '" & txtTransNo & "'"
gc_dbcon.Execute ls_sql
ls_sql = "Delete from gl_trans where compcode = '" & Gs_compcode & "' and vchrtype = 'PRV' and Voucher_no = '" & txtTransNo & "'"
gc_dbcon.Execute ls_sql

End Sub
Private Sub DeleteBankVoucher()
ls_sql = "Delete from gl_ref where compcode = '" & Gs_compcode & "' and vchrtype = 'BPP' and Voucher_no = '" & txtTransNo & "'"
gc_dbcon.Execute ls_sql
ls_sql = "Delete from gl_trans where compcode = '" & Gs_compcode & "' and vchrtype = 'BPP' and Voucher_no = '" & txtTransNo & "'"
gc_dbcon.Execute ls_sql
End Sub
Private Sub DeleteCashVoucher()
ls_sql = "Delete from gl_ref where compcode = '" & Gs_compcode & "' and vchrtype = 'CPP' and Voucher_no = '" & txtTransNo & "'"
gc_dbcon.Execute ls_sql
ls_sql = "Delete from gl_trans where compcode = '" & Gs_compcode & "' and vchrtype = 'CPP' and Voucher_no = '" & txtTransNo & "'"
gc_dbcon.Execute ls_sql
End Sub

Private Sub txtClientCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtClientCode <> "" Then
    txtClientCode = DoPad(txtClientCode, 6)
    ls_sql = "Select SupplierCode,Description from IC_Supplier where compcode = '" & Gs_compcode & "' and SupplierCode = '" & txtClientCode & "' "
    PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
     If PR_Dumy.EOF Then
            Call MsgBox("Client code not found !!!", vbCritical)
            txtClientCode = ""
            txtclientdesc = ""
            txtClientCode.SetFocus
     Else
                txtclientdesc = PR_Dumy("description")
                cmdGenerate.SetFocus
     End If
          PR_Dumy.Close



ElseIf KeyCode = vbKeyReturn And txtClientCode = "" Then
        Command5_Click
End If

End Sub


Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtTransNo) <> "" And KeyCode = vbKeyReturn Then
        txtTransNo = DoPad(txtTransNo, 10)
    
    If txtstatus.ListIndex = 0 Then
    
        ls_sql = "Select PO_POGRN.TransCode, PO_POGRN.Transdate,PO_POGRN.accountcode,IC_Supplier.Description from  PO_POGRN left join IC_Supplier on PO_POGRN.AccountCode = IC_Supplier.SupplierCode "
        ls_sql = ls_sql & " where PO_POGRN.Compcode ='" & Gs_compcode & "'   and PO_POGRN.transcode = '" & txtTransNo & "' "
        
        PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Document Code not found !!!", vbCritical)
            txtTransNo = ""
            txtdesc = ""
            txtTransNo.SetFocus
        Else
            txtdesc = PR_Dumy("Transdate")
            txtvaluedate = PR_Dumy("Transdate")
            txtoldclientcode = Trim(PR_Dumy("Accountcode") & "")
            lbloldclient.Caption = Trim(PR_Dumy("Description") & "")
            txtremarks.SetFocus
        End If
        PR_Dumy.Close
        
    ElseIf txtstatus.ListIndex = 1 Then
    
            ls_sql = "Select PO_POGRNReturn.TransCode, PO_POGRNReturn.Transdate,PO_POGRNReturn.accountcode,IC_Supplier.Description from  PO_POGRNReturn left join IC_Supplier on PO_POGRNReturn.AccountCode = IC_Supplier.SupplierCode "
            ls_sql = ls_sql & " where PO_POGRNReturn.Compcode ='" & Gs_compcode & "' and PO_POGRNReturn.transcode = '" & txtTransNo & "' "
            
            PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Document Code not found !!!", vbCritical)
                txtTransNo = ""
                txtdesc = ""
                txtTransNo.SetFocus
            Else
                txtdesc = PR_Dumy("Transdate")
                txtvaluedate = PR_Dumy("Transdate")
                txtoldclientcode = Trim(PR_Dumy("Accountcode") & "")
                lbloldclient.Caption = Trim(PR_Dumy("Description") & "")
                txtremarks.SetFocus
            End If
            PR_Dumy.Close
    ElseIf txtstatus.ListIndex = 2 Then
    
            ls_sql = "Select * from  PO_PayableMaster "
            ls_sql = ls_sql & " where Compcode ='" & Gs_compcode & "' and transcode = '" & txtTransNo & "' "
            
            PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Document Code not found !!!", vbCritical)
                txtTransNo = ""
                txtdesc = ""
                txtTransNo.SetFocus
            Else
                txtdesc = PR_Dumy("Transdate")
                txtvaluedate = PR_Dumy("Transdate")
                txtremarks.SetFocus
            End If
            PR_Dumy.Close
    ElseIf txtstatus.ListIndex = 3 Then
    
            ls_sql = "Select * from  PO_PayableCashMaster "
            ls_sql = ls_sql & " where Compcode ='" & Gs_compcode & "' and transcode = '" & txtTransNo & "' "
            
            PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Document Code not found !!!", vbCritical)
                txtTransNo = ""
                txtdesc = ""
                txtTransNo.SetFocus
            Else
                txtdesc = PR_Dumy("Transdate")
                txtvaluedate = PR_Dumy("Transdate")
                txtremarks.SetFocus
            End If
            PR_Dumy.Close
    
    
    End If

ElseIf Trim(txtTransNo) = "" And KeyCode = vbKeyReturn Then
        txtTransNo = ""
        txtdesc = ""
        Command5_Click
End If

End Sub

