VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmicreport12 
   Caption         =   "Stock Zero Form"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports12.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   315
      Left            =   2880
      Picture         =   "frmicreports12.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1080
      Width           =   315
   End
   Begin VB.TextBox txtItemDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4185
   End
   Begin VB.TextBox txtSubDeptDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   720
      Width           =   4185
   End
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   2880
      Picture         =   "frmicreports12.frx":047C
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   720
      Width           =   315
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4650
      Width           =   7635
      _ExtentX        =   13467
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
   Begin VB.Frame Frame1 
      Height          =   3510
      Left            =   30
      TabIndex        =   6
      Top             =   -60
      Width           =   7560
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   2925
         Left            =   105
         TabIndex        =   8
         Top             =   120
         Width           =   7425
         Begin VB.TextBox txtItemCode 
            Height          =   315
            Left            =   1680
            TabIndex        =   23
            Top             =   960
            Width           =   1020
         End
         Begin VB.TextBox txtSubDeptCode 
            Height          =   315
            Left            =   1680
            TabIndex        =   22
            Top             =   600
            Width           =   1020
         End
         Begin VB.TextBox txtadjno 
            Height          =   315
            Left            =   1680
            TabIndex        =   14
            Top             =   2520
            Width           =   1845
         End
         Begin VB.TextBox txtselectedcode 
            Height          =   315
            Left            =   1680
            TabIndex        =   0
            Top             =   195
            Width           =   990
         End
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   195
            Width           =   4185
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2760
            Picture         =   "frmicreports12.frx":05EE
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   195
            Width           =   315
         End
         Begin VB.ComboBox txtStoreType 
            Height          =   330
            ItemData        =   "frmicreports12.frx":0760
            Left            =   1680
            List            =   "frmicreports12.frx":076A
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1560
            Width           =   1860
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1680
            TabIndex        =   1
            Top             =   2040
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   16580609
            CurrentDate     =   37293
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sub Department :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Item Code :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   360
            TabIndex        =   24
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "For Repost Stock"
            Height          =   285
            Left            =   3720
            TabIndex        =   17
            Top             =   2520
            Width           =   1245
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   30
            Left            =   3615
            TabIndex        =   16
            Top             =   1500
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Adj Note # :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   720
            TabIndex        =   15
            Top             =   2520
            Width           =   840
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Department Code :"
            Height          =   210
            Left            =   195
            TabIndex        =   13
            Top             =   225
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Adj Entry  Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   360
            TabIndex        =   10
            Top             =   2160
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Store Type :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   600
            TabIndex        =   9
            Top             =   1560
            Width           =   885
         End
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   5160
         Top             =   1185
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   6480
         TabIndex        =   4
         Top             =   3120
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Stock Zero"
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
         Left            =   5400
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   3120
         Width           =   1035
      End
      Begin VB.TextBox txtVchrDesc 
         Height          =   315
         Left            =   435
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label txtselectiveaccount 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3840
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Label txtselectiveitem 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   600
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   2520
   End
End
Attribute VB_Name = "frmicreport12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Public codeid As String
Public Reporttype As String
Dim ls_sql As String




Private Sub Check1_Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from IC_InventoryAdjMaster where compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close

End Function

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr

If txtselectedcode = "" Then
Call MsgBox("Select Department Code", vbCritical)
txtselectedcode.SetFocus
Exit Sub
End If

If txtselectedcode = "" Then
Call MsgBox("Select Department Code", vbCritical)
txtselectedcode.SetFocus
Exit Sub
End If

gc_dbcon.Execute "UPDATE IC_Item SET   StockG = 0 ,StockS = 0"
gc_dbcon.Execute "UPDATE IC_Item SET   StockG = StockSummary.Qty FROM   StockSummary INNER JOIN   IC_Item ON StockSummary.ItemCode = IC_Item.ItemCode WHERE (StockSummary.siteid = 1)"
gc_dbcon.Execute "UPDATE IC_Item SET   StockS = StockSummary.Qty FROM   StockSummary INNER JOIN   IC_Item ON StockSummary.ItemCode = IC_Item.ItemCode WHERE (StockSummary.siteid = 2)"
Call MsgBox("Stock Successfully Update !!!", vbInformation)


MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."

 Call ChkTempTables("Tmp_Stock", True)


If txtadjno <> "" Then
 gc_dbcon.Execute "DELETE FROM IC_InventoryAdjMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtadjno) & "'"
 gc_dbcon.Execute "DELETE FROM IC_InventoryAdjDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtadjno) & "'"
End If



ls_sql = "SELECT a.compcode, a.ItemCode,0 as Invtype, (SUM(a.PQty) + SUM(a.SRQty)) - (SUM(a.SQty) + SUM(a.IQTY) + SUM(a.RPQty))"
ls_sql = ls_sql & " AS OpQty into Tmp_Stock FROM         StockLedgerDetail a left outer join ic_item b on  a.compcode = b.compcode and  a.itemcode = b.itemcode"
ls_sql = ls_sql & " WHERE     (a.compcode = '" & Gs_compcode & "' )  AND (a.siteid = " & txtStoreType.ListIndex + 1 & ")  and b.catcode =  '" & txtselectedcode & "'  group by a.compcode, a.ItemCode"
gc_dbcon.Execute ls_sql


ls_sql = "delete from Tmp_Stock where OpQty = 0"
gc_dbcon.Execute ls_sql

If txtadjno = "" Then
txtadjno = maxtranscode
Else

 gc_dbcon.Execute "DELETE FROM IC_InventoryAdjMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtadjno) & "'"
 gc_dbcon.Execute "DELETE FROM IC_InventoryAdjDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtadjno) & "'"

End If



ls_sql = "INSERT into IC_InventoryAdjMaster( Compcode,branchcode, TransCode,   TransDate, AccountCode, SiteID, BinID, Remarks,vcode,acode,userid,adddate,addtime,InvType,adjsiteid)"
ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtadjno) & "','" & Format(dtpfrom, "YYYY/MM/DD") & "','000001','001','001','Auto Adjustment in Inventory','001','001','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "',1," & txtStoreType.ListIndex + 1 & "  )"
gc_dbcon.Execute ls_sql



ls_sql = "insert into IC_InventoryAdjDetail (Compcode, BranchCode, TransCode, customcode, ItemCode, InvType, Quantity, ItemRate, AvgRate, Amount, Remarks)"

ls_sql = ls_sql & "  SELECT  a.compcode, '001' as Brcode ,'" & txtadjno & "',b.customcode,a.ItemCode,0 as Invtype, a.OpQty,b.salecost,b.salecost,b.salecost* a.opqty as Amount,'H' as Remarks"
ls_sql = ls_sql & "  From Tmp_Stock a left outer join ic_item b on  a.compcode = b.compcode and  a.itemcode = b.itemcode where a.opqty > 0"
gc_dbcon.Execute ls_sql


txtadjno = maxtranscode

ls_sql = "INSERT into IC_InventoryAdjMaster( Compcode,branchcode, TransCode,   TransDate, AccountCode, SiteID, BinID, Remarks,vcode,acode,userid,adddate,addtime,InvType,adjsiteid)"
ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtadjno) & "','" & Format(dtpfrom, "YYYY/MM/DD") & "','000001','001','001','Auto Adjustment in Inventory','001','001','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "',0," & txtStoreType.ListIndex + 1 & "  )"
gc_dbcon.Execute ls_sql



ls_sql = "insert into IC_InventoryAdjDetail (Compcode, BranchCode, TransCode, customcode, ItemCode, InvType, Quantity, ItemRate, AvgRate, Amount, Remarks)"

ls_sql = ls_sql & "  SELECT  a.compcode, '001' as Brcode ,'" & txtadjno & "',b.customcode,a.ItemCode,0 as Invtype, (-1)*a.OpQty,b.salecost,b.salecost,b.salecost* (-1)*(a.opqty) as Amount,'H' as Remarks"
ls_sql = ls_sql & "  From Tmp_Stock a left outer join ic_item b on  a.compcode = b.compcode and  a.itemcode = b.itemcode where a.opqty < 0"
gc_dbcon.Execute ls_sql


'ls_sql = " update IC_InventoryAdjDetail set invtype = 1  where transcode = '" & txtadjno & "'  and quantity > 0"
'gc_dbcon.Execute ls_sql

'ls_sql = " update IC_InventoryAdjDetail set quantity = (-1)*quantity  where transcode = '" & txtadjno & "' and quantity < 0"
'gc_dbcon.Execute ls_sql

'MDIForm1.StatusBar1.Panels(7).Text = ""

Call MsgBox("Successfully Updated  Adjustment Note # " + txtadjno, vbInformation)

Exit Sub

LocalErr:
Call MsgBox(Err.Description, vbCritical)
On Error GoTo 0
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectedcode
    Set PO_DESC = txtdesc
    Gs_SQL = "SELECT CatCode, Description  FROM IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Department"
    MyLookupOLDB.Show 1
    If txtselectedcode <> "" Then Call txtselectedcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSubDeptCode
    Set PO_DESC = txtSubDeptDesc
    Gs_SQL = "SELECT Classcode, Description  FROM IC_ItemClass"
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and DeptCode='" & txtselectedcode.Text & "'"
    MyLookupOLDB.Caption = "Sub Department"
    MyLookupOLDB.Show 1
    'If txtSubDeptCode <> "" Then Call txtSubDeptCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtItemCode
    Set PO_DESC = txtItemDesc
    Gs_SQL = "SELECT itemcode, Description  FROM IC_Item"
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and CatCode='" & txtselectedcode.Text & "' and ClassID='" & txtSubDeptCode.Text & "'"
    MyLookupOLDB.Caption = "Item Information"
    MyLookupOLDB.Show 1
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   txtStoreType.SetFocus
End If
End Sub

Private Sub Form_Load()
  dtpfrom = Date
  txtStoreType = "GODOWN"
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub txtadjno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtadjno <> "" Then txtadjno = DoPad(txtadjno, 10)
End Sub

Private Sub txtadjno_LostFocus()
If txtadjno <> "" Then txtadjno = DoPad(txtadjno, 10)
End Sub



Private Sub txtselectedcode_Change()
If txtselectedcode = "" Then
txtdesc = ""
End If
End Sub


Private Sub txtselectedcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtselectedcode <> "" And KeyCode = vbKeyReturn Then
    txtselectedcode = DoPad(txtselectedcode, 3)
    ls_sql = "Select Catcode,Description from IC_ItemCategory where compcode = '" & Gs_compcode & "' and Catcode = '" & txtselectedcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Department Code not found", vbCritical)
            Else
                txtdesc = pr_dumy("description")
                txtStoreType.SetFocus
            End If
         pr_dumy.Close

End If
End Sub

Private Sub txtStoreType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpfrom.SetFocus
End Sub
