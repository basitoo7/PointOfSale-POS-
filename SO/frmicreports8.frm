VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmicreport8 
   Caption         =   "Sale Register Report"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports8.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3525
      Width           =   6735
      _ExtentX        =   11880
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
      Height          =   3570
      Left            =   30
      TabIndex        =   5
      Top             =   -90
      Width           =   6675
      Begin VB.CheckBox chksitemonly 
         Caption         =   "Sale Items Report Only"
         Height          =   270
         Left            =   150
         TabIndex        =   30
         Top             =   3150
         Width           =   2385
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   5550
         TabIndex        =   4
         Top             =   3150
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
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
         Left            =   4470
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   3150
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   -15
         Top             =   3435
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
      Begin VB.TextBox txtVchrDesc 
         Height          =   315
         Left            =   15
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   375
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   3075
         Left            =   0
         TabIndex        =   6
         Top             =   30
         Width           =   6630
         Begin VB.CheckBox chkitemsaledeptwise 
            Caption         =   "Item Sale Quantity Dept Wise."
            Height          =   270
            Left            =   3630
            TabIndex        =   31
            Top             =   285
            Width           =   2820
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   3000
            Picture         =   "frmicreports8.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2535
            Width           =   315
         End
         Begin VB.TextBox txtitemdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3345
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   2535
            Width           =   3210
         End
         Begin VB.TextBox txtitemcode 
            Height          =   315
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   22
            Top             =   2550
            Width           =   960
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2610
            Picture         =   "frmicreports8.frx":047C
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1005
            Width           =   315
         End
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2970
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1005
            Width           =   3585
         End
         Begin VB.TextBox txtdeptcode 
            Height          =   315
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   19
            Top             =   1005
            Width           =   555
         End
         Begin VB.TextBox txtcatcode 
            Height          =   315
            Left            =   2040
            TabIndex        =   18
            Top             =   1380
            Width           =   555
         End
         Begin VB.TextBox txtcatdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2970
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1380
            Width           =   3585
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2625
            Picture         =   "frmicreports8.frx":05EE
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1395
            Width           =   315
         End
         Begin VB.TextBox txtsubcatcode 
            Height          =   315
            Left            =   2040
            TabIndex        =   15
            Top             =   1755
            Width           =   555
         End
         Begin VB.TextBox txtsubcatdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2970
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1755
            Width           =   3585
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2610
            Picture         =   "frmicreports8.frx":0760
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1755
            Width           =   315
         End
         Begin VB.TextBox txtSuppliercode 
            Height          =   315
            Left            =   2040
            TabIndex        =   12
            Top             =   2160
            Width           =   945
         End
         Begin VB.TextBox txtSupplierdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3345
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2160
            Width           =   3210
         End
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   3000
            Picture         =   "frmicreports8.frx":08D2
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2160
            Width           =   315
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   2040
            TabIndex        =   1
            Top             =   255
            Width           =   1545
            _ExtentX        =   2725
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
            CustomFormat    =   "MMM d yyyy hh:mmtt"
            Format          =   64225281
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   2040
            TabIndex        =   2
            Top             =   615
            Width           =   1560
            _ExtentX        =   2752
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
            CustomFormat    =   "MMM d yyyy hh:mmtt"
            Format          =   64225281
            CurrentDate     =   37293
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Item Code :"
            Height          =   210
            Left            =   1215
            TabIndex        =   29
            Top             =   2565
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Department Code :"
            Height          =   210
            Left            =   645
            TabIndex        =   28
            Top             =   1035
            Width           =   1335
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Category Code :"
            Height          =   210
            Left            =   810
            TabIndex        =   27
            Top             =   1410
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sub Cat Code :"
            Height          =   210
            Left            =   900
            TabIndex        =   26
            Top             =   1785
            Width           =   1080
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Supplier Code :"
            Height          =   210
            Left            =   885
            TabIndex        =   25
            Top             =   2190
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   1170
            TabIndex        =   8
            Top             =   645
            Width           =   825
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1170
            TabIndex        =   7
            Top             =   270
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frmicreport8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Dumy As New Recordset
Dim PR_Branch As New Recordset
Public codeid As String
Dim ls_sql As String
Dim ls_branchdesc As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."

Call ChkTempTables("Tmp_SaleReport1", True)
Call ChkTempTables("Tmp_SaleReport", True)

ls_sql = "SELECT itemcode, SUM(Amount)  as Saleamount , SUM(DiscAmount) AS Discamount, SUM(SaleQty) AS SaleQty, SUM(SaleReturn) AS SaleReturn, SUM(ReturnQty) AS ReturnQty,0 as opqty into Tmp_SaleReport1 From SaleReport"
ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'  "
ls_sql = ls_sql & " and convert(varchar,transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
ls_sql = ls_sql & " and convert(varchar,transdate,111) <= '" & Format(dtpto.Value, "YYYY/MM/DD") & "' "
ls_sql = ls_sql & " group by itemcode "
ls_sql = ls_sql & " Union all"
            
ls_sql = ls_sql & " SELECT ItemCode, 0 as samount,0 as Discamount,0 as saleqty, 0 as salereturn, 0 as SRQty,(sum(PQty)+sum(SRQty))- (sum(SQty)+sum(IQTY)+sum(RPQty)) as OpQty  From StockLedgerDetail"
ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'"
         
ls_sql = ls_sql & " and transdate < '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by itemcode"
       
gc_dbcon.Execute ls_sql

ls_sql = "SELECT itemcode, SUM(Saleamount)  as Saleamount , SUM(DiscAmount) AS Discamount, SUM(SaleQty) AS SaleQty, SUM(SaleReturn) AS SaleReturn, SUM(ReturnQty) AS ReturnQty,sum(opqty) as opqty into Tmp_SaleReport From Tmp_SaleReport1 group by itemcode"
gc_dbcon.Execute ls_sql


   With crrpt
        If chkitemsaledeptwise.Value = 1 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportCateSummary.RPT"
         chksitemonly.Value = 1
        Else
        .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportManusummary.RPT"
        End If
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = '" & Me.Caption & "[" & txtdesc & "]" & "[" & txtcatdesc & "]" & "[" & txtsubcatdesc & "]" & "'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
       
        .SelectionFormula = "{Ic_item.CompCode} = '" & Gs_compcode & "'"
      
       If chksitemonly.Value = 1 Then
       .SelectionFormula = .SelectionFormula & " and {Tmp_SaleReport.Saleamount} <> 0"
       End If
      
      
       If txtdeptcode <> "" Then
       .SelectionFormula = .SelectionFormula & " and {Ic_item.Catcode} = '" & txtdeptcode & "'"
       End If
      
      
       If txtcatcode <> "" Then
       .SelectionFormula = .SelectionFormula & " and {Ic_item.classid} = '" & txtcatcode & "'"
       End If
       
       
        If txtsubcatcode <> "" Then
       .SelectionFormula = .SelectionFormula & " and {Ic_item.packcode} = '" & txtsubcatcode & "'"
       End If
       
       If txtSuppliercode <> "" Then
       .SelectionFormula = .SelectionFormula & " and {Ic_item.manucode} = '" & txtSuppliercode & "'"
       End If
      
        
       If txtitemcode <> "" Then
       .SelectionFormula = .SelectionFormula & " and {Ic_item.itemcode} = '" & txtitemcode & "'"
       End If
        
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
   
   
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
Call MsgBox(Err.Description, vbCritical)

On Error GoTo 0
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtdeptcode
    Set PO_DESC = txtdesc
    Gs_SQL = "Select CatCode,   Description from IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    
    If txtdeptcode <> "" Then Call txtdeptcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcatcode
    Set PO_DESC = txtcatdesc
    Gs_SQL = "Select ClassCode,   Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "'"
    MyLookupOLDB.Caption = "Categories"
    MyLookupOLDB.Show 1
    
    If txtcatcode <> "" Then Call txtcatcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command3_Click()
  Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsubcatcode
    Set PO_DESC = txtsubcatdesc
    Gs_SQL = "Select PackCode,   Description from IC_ItemPacking "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and subcode = '" & txtcatcode & "' and deptcode = '" & txtdeptcode & "' "
    MyLookupOLDB.Caption = "Sub Categories"
    MyLookupOLDB.Show 1
    
    If txtsubcatcode <> "" Then Call txtsubcatcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command4_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSuppliercode
    Set PO_DESC = txtSupplierdesc
    Gs_SQL = "Select SupplierCode,   Description from IC_Supplier "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Supplier"
    MyLookupOLDB.Show 1

    
    If txtSuppliercode <> "" Then Call txtSuppliercode_KeyDown(vbKeyReturn, vbKeyShift)


End Sub

Private Sub Command5_Click()
   Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtitemcode
    Set PO_DESC = txtitemdesc
    
    Gs_SQL = "SELECT itemCode, Description  FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
    
    If txtitemcode <> "" Then Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub txtcatcode_Change()
If txtcatcode = "" Then
txtcatdesc = ""
End If
End Sub

Private Sub txtcatcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtcatcode) <> "" And KeyCode = vbKeyReturn Then
        txtcatcode = DoPad(txtcatcode, 3)
       If PR_Dumy.State = 1 Then PR_Dumy.Close
        PR_Dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtcatcode & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Category Code not found !!!", vbCritical)
            txtcatcode = ""
            txtcatdesc = ""
            txtcatcode.SetFocus
        Else
            txtcatdesc = PR_Dumy("Description")
             If txtsubcatcode.Enabled Then txtsubcatcode.SetFocus
           
        End If
        PR_Dumy.Close
        
ElseIf Trim(txtcatcode) = "" And KeyCode = vbKeyReturn Then
        txtcatcode = ""
        txtcatdesc = ""
        Command2_Click
End If
End Sub

Private Sub txtdeptcode_Change()
If txtdeptcode = "" Then
txtdesc = ""
End If
End Sub

Private Sub txtdeptcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtdeptcode <> "" And KeyCode = vbKeyReturn Then
    txtdeptcode = DoPad(txtdeptcode, txtdeptcode.MaxLength)
    ls_sql = "Select Catcode,Description from IC_ItemCategory where compcode = '" & Gs_compcode & "' and Catcode = '" & txtdeptcode & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Department Code not found", vbCritical)
            Else
                txtdesc = PR_Dumy("description")
               txtcatcode.SetFocus
            End If
         PR_Dumy.Close
ElseIf txtdeptcode = "" And KeyCode = vbKeyReturn Then
Command1_Click
End If
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub


Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtdeptcode.SetFocus
End If
End Sub

Private Sub Form_Load()
  dtpfrom = Date
  dtpto = Date
 
End Sub


Private Sub txtItemcode_Change()
If txtitemcode = "" Then
txtitemdesc = ""
End If

End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtitemcode <> "" And KeyCode = vbKeyReturn Then
    txtitemcode = DoPad(txtitemcode, txtitemcode.MaxLength)
    ls_sql = "Select itemcode,Description from IC_Item where compcode = '" & Gs_compcode & "' and itemcode = '" & txtitemcode & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Item Code not found", vbCritical)
            Else
                txtitemdesc = PR_Dumy("description")
                 dtpfrom.SetFocus
            End If
         PR_Dumy.Close
ElseIf txtitemcode = "" And KeyCode = vbKeyReturn Then
    Command5_Click
End If
End Sub

Private Sub txtsubcatcode_Change()
If txtsubcatcode = "" Then
txtsubcatdesc = ""
End If
End Sub

Private Sub txtsubcatcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtsubcatcode) <> "" And KeyCode = vbKeyReturn Then
        txtsubcatcode.Text = DoPad(txtsubcatcode.Text, 3)
        PR_Dumy.Open "Select * from IC_ItemPacking where Packcode = '" & txtsubcatcode & "'  and subcode = '" & txtcatcode & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Sub Category code not found !!!", vbCritical)
            txtsubcatcode = ""
            txtsubcatdesc = ""
            txtsubcatcode.SetFocus
        Else
            txtsubcatdesc = PR_Dumy("Description")
            If txtSuppliercode.Enabled Then txtSuppliercode.SetFocus
            
        End If
        PR_Dumy.Close
        
ElseIf Trim(txtsubcatcode) = "" And KeyCode = vbKeyReturn Then
        txtsubcatcode = ""
        txtsubcatdesc = ""
        Command3_Click
End If

End Sub

Private Sub txtSuppliercode_Change()
If txtSuppliercode = "" Then
txtSupplierdesc = ""
End If
End Sub

Private Sub txtSuppliercode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtSuppliercode) <> "" And KeyCode = vbKeyReturn Then
        txtSuppliercode.Text = DoPad(txtSuppliercode.Text, 6)
        PR_Dumy.Open "Select * from IC_Supplier where Suppliercode = '" & txtSuppliercode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Supplier code not found !!!", vbCritical)
            txtSuppliercode = ""
            txtSupplierdesc = ""
            txtSuppliercode.SetFocus
        Else
            txtSupplierdesc = PR_Dumy("Description")
            If txtitemcode.Enabled Then txtitemcode.SetFocus
            
        End If
        PR_Dumy.Close
ElseIf Trim(txtSuppliercode) = "" And KeyCode = vbKeyReturn Then
        txtSuppliercode = ""
        txtSupplierdesc = ""
        Command4_Click
End If

End Sub

