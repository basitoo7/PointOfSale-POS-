VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmicreportCasher 
   Caption         =   "Sale Report"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreportsCasher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4890
      Width           =   8055
      _ExtentX        =   14208
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
      Height          =   4950
      Left            =   60
      TabIndex        =   3
      Top             =   -60
      Width           =   7935
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   1875
         Left            =   0
         TabIndex        =   4
         Top             =   30
         Width           =   7920
         Begin VB.ComboBox txtcasher1 
            Height          =   330
            ItemData        =   "frmicreportsCasher.frx":030A
            Left            =   1485
            List            =   "frmicreportsCasher.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1425
            Width           =   2505
         End
         Begin VB.CheckBox Chktime 
            Caption         =   "Sale With Time"
            Height          =   210
            Left            =   3900
            TabIndex        =   15
            Top             =   165
            Width           =   1650
         End
         Begin VB.ComboBox txtcasher 
            Height          =   330
            ItemData        =   "frmicreportsCasher.frx":030E
            Left            =   1485
            List            =   "frmicreportsCasher.frx":0310
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1440
            Width           =   2505
         End
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   1245
            Left            =   3855
            TabIndex        =   9
            Top             =   120
            Width           =   4005
            Begin MSComCtl2.DTPicker DTPtimefrom 
               Height          =   315
               Left            =   1545
               TabIndex        =   10
               Top             =   315
               Width           =   1350
               _ExtentX        =   2381
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
               CustomFormat    =   "HH:mm:ss"
               Format          =   101646339
               CurrentDate     =   37293
            End
            Begin MSComCtl2.DTPicker DTPtimeto 
               Height          =   315
               Left            =   1545
               TabIndex        =   11
               Top             =   690
               Width           =   1365
               _ExtentX        =   2408
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
               CustomFormat    =   "HH:mm:ss"
               Format          =   94699523
               CurrentDate     =   37293
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "To Time :"
               Height          =   210
               Left            =   675
               TabIndex        =   13
               Top             =   705
               Width           =   825
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "From Time :"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   675
               TabIndex        =   12
               Top             =   330
               Width           =   825
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1245
            Left            =   45
            TabIndex        =   31
            Top             =   120
            Width           =   3750
            Begin MSComCtl2.DTPicker dtpfrom 
               Height          =   315
               Left            =   1410
               TabIndex        =   32
               Top             =   270
               Width           =   2085
               _ExtentX        =   3678
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
               Format          =   103088129
               CurrentDate     =   37293
            End
            Begin MSComCtl2.DTPicker dtpto 
               Height          =   315
               Left            =   1410
               TabIndex        =   33
               Top             =   645
               Width           =   2100
               _ExtentX        =   3704
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
               Format          =   103088129
               CurrentDate     =   37293
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "To Date :"
               Height          =   210
               Left            =   705
               TabIndex        =   35
               Top             =   675
               Width           =   645
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "From Date :"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   525
               TabIndex        =   34
               Top             =   300
               Width           =   825
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Casher :"
            Height          =   210
            Left            =   750
            TabIndex        =   37
            Top             =   1455
            Width           =   615
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save "
         Height          =   330
         Left            =   6705
         TabIndex        =   28
         Top             =   3765
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Update"
         Height          =   330
         Left            =   5655
         TabIndex        =   27
         Top             =   3765
         Width           =   1035
      End
      Begin VB.ComboBox txtfields 
         Height          =   330
         ItemData        =   "frmicreportsCasher.frx":0312
         Left            =   945
         List            =   "frmicreportsCasher.frx":0319
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2415
         Width           =   2565
      End
      Begin VB.TextBox txtvalueadd 
         Height          =   870
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2835
         Width           =   7725
      End
      Begin VB.ComboBox txtEq 
         Height          =   330
         ItemData        =   "frmicreportsCasher.frx":032E
         Left            =   3495
         List            =   "frmicreportsCasher.frx":033E
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2415
         Width           =   960
      End
      Begin VB.ComboBox txtandor 
         Height          =   330
         ItemData        =   "frmicreportsCasher.frx":0351
         Left            =   30
         List            =   "frmicreportsCasher.frx":035B
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2400
         Width           =   870
      End
      Begin VB.TextBox txtvalue 
         Height          =   315
         Left            =   4500
         TabIndex        =   21
         Top             =   2430
         Width           =   2250
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   300
         Left            =   6765
         TabIndex        =   20
         Top             =   2430
         Width           =   990
      End
      Begin VB.ComboBox txtcustomReport 
         Height          =   330
         ItemData        =   "frmicreportsCasher.frx":0368
         Left            =   1470
         List            =   "frmicreportsCasher.frx":036A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1995
         Width           =   4785
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   6765
         TabIndex        =   18
         Top             =   1995
         Width           =   990
      End
      Begin VB.CheckBox chkcredit 
         Caption         =   "Credit"
         Height          =   345
         Left            =   6735
         TabIndex        =   17
         Top             =   4140
         Width           =   1005
      End
      Begin VB.CheckBox Chkcash 
         Caption         =   "Cash"
         Height          =   345
         Left            =   5835
         TabIndex        =   16
         Top             =   4140
         Width           =   1005
      End
      Begin VB.CheckBox Chkwinvoicetotal 
         Caption         =   "With Invoice Total"
         Height          =   225
         Left            =   45
         TabIndex        =   14
         Top             =   4200
         Width           =   1770
      End
      Begin VB.ComboBox txtreporttye 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmicreportsCasher.frx":036C
         Left            =   1500
         List            =   "frmicreportsCasher.frx":0391
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3795
         Width           =   4020
      End
      Begin VB.CheckBox ChkSummary 
         Caption         =   "Summary Only :"
         Height          =   225
         Left            =   45
         TabIndex        =   6
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   6705
         TabIndex        =   2
         Top             =   4545
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
         Left            =   5655
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Top             =   4545
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   5760
         Top             =   1575
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
         Left            =   4140
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox txtcustomReport1 
         Height          =   330
         ItemData        =   "frmicreportsCasher.frx":04D9
         Left            =   1515
         List            =   "frmicreportsCasher.frx":04DB
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1995
         Visible         =   0   'False
         Width           =   4965
      End
      Begin VB.ComboBox txtid 
         Height          =   330
         ItemData        =   "frmicreportsCasher.frx":04DD
         Left            =   1500
         List            =   "frmicreportsCasher.frx":04DF
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1995
         Visible         =   0   'False
         Width           =   5400
      End
      Begin VB.Label Label4 
         Caption         =   "Custom Reports :"
         Height          =   300
         Left            =   135
         TabIndex        =   26
         Top             =   2010
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmicreportCasher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Dim PR_Branch As New Recordset
Public codeid As String
Dim ls_sql As String
Dim ls_branchdesc As String
Private Sub Command2_Click()
Dim res
res = InputBox("Name of the Report", "Custom Reort")
If res <> "" Then
gc_dbcon.Execute "Insert  into so_customreport (reportname,parameter,reportid) values('" & res & "','" & txtvalueadd & "',2) "
Loadcustomreport
End If

End Sub

Private Sub Command5_Click()
UpdateCostofNonAvgRate
Call MsgBox("Successfully Updated", vbInformation)
End Sub

Private Sub txtcustomReport_Click()
txtcustomReport1.ListIndex = txtcustomReport.ListIndex
txtvalueadd = txtcustomReport1.Text
txtid.ListIndex = txtcustomReport1.ListIndex
End Sub

Private Sub txtcustomReport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtcustomReport_Click
End If
End Sub
Private Sub Command3_Click()
Dim ls_id
ls_id = txtid.Text
gc_dbcon.Execute "delete from so_Customreport where id = " & ls_id & " and reportid = 2 "
Loadcustomreport
End Sub

Private Sub Command4_Click()
On Error GoTo localerror
Dim ls_id
ls_id = txtid.Text
gc_dbcon.Execute "delete from so_Customreport where id = " & ls_id & " and reportid = 2 "
gc_dbcon.Execute "Insert  into so_customreport (reportname,parameter,reportid) values('" & txtcustomReport.Text & "','" & txtvalueadd & "',2) "
Loadcustomreport
Exit Sub
localerror:

End Sub
Private Sub Loadcustomreport()
Dim pr_CustomReport As New Recordset
   txtcustomReport.Clear
   txtcustomReport1.Clear
   txtid.Clear

pr_CustomReport.Open "SELECT *  from So_customreport where reportid = 2", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_CustomReport.EOF Then
Do While Not pr_CustomReport.EOF
   txtcustomReport.AddItem pr_CustomReport("reportname")
   txtcustomReport1.AddItem pr_CustomReport("parameter")
   txtid.AddItem pr_CustomReport("ID")
pr_CustomReport.MoveNext
Loop
End If
pr_CustomReport.Close
End Sub

Private Sub Check1_Click()

End Sub

Private Sub ChkSummary_Click()
If ChkSummary.Value = 1 Then
    txtreporttye.Enabled = True
Else
    txtreporttye.Enabled = False
End If
End Sub

Private Sub Chktime_Click()
If Chktime.Value = 1 Then
Frame2.Enabled = True
    DTPtimefrom.SetFocus
Else
    Frame2.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
'On Error GoTo LocalErr

MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."

'WebBrowser1.Navigate ("http://api.bizsms.pk/api-send-branded-sms.aspx?username=makori@bizsms.pk&pass=mq3t897&text=" & Trim(txtSearch) & "&masking=Rahat Store&destinationnum=" & Trim(txtUrlLoc.Text) & "&language=English")

'******************** Start Amjid sb Reports ****************

'If txtreporttye.ListIndex = 9 Then
'Call ChkTempTables("Tmp_SaleReport", True)

'ls_sql = "SELECT itemcode, (Sum(Amount) - Sum(DiscAmount)) AS Saleamount, Sum(SaleQty) AS SaleQty, Sum(SaleReturn) AS SaleReturn, Sum(ReturnQty) AS ReturnQty,Sum(costamount) as Costamount, 0 as PSaleAmount , 0 as PSaleQty,0 as  PSaleReturn ,0 as PReturnQty, 0 as  Pcostamount into Tmp_SaleReport From Ecounts.dbo.SaleReport where TransDate >='2017/04/01' and TransDate <='2017/04/30'"

'ls_sql = "SELECT itemcode, SUM(Amount) - SUM(DiscAmount) AS Saleamount, SUM(SaleQty) AS SaleQty, SUM(SaleReturn) AS SaleReturn, SUM(ReturnQty) AS ReturnQty,SUM(costamount) as Costamount,0  into Tmp_SaleReport From SaleReport"
'ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'  "

'ls_sql = ls_sql & " and convert(varchar,transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
'ls_sql = ls_sql & " and convert(varchar,transdate,111) <= '" & Format(DTPTo.Value, "YYYY/MM/DD") & "' "

'ls_sql = ls_sql & "Union All"
'ls_sql = " SELECT itemcode, 0 as SaleAmount, 0 as SaleQty, 0 as SaleReturn, 0 as ReturnQty, 0 as CostAmount, (Sum(Amount) - Sum(DiscAmount)) AS PSaleamount, Sum(SaleQty) AS PSaleQty, Sum(SaleReturn) AS PSaleReturn, Sum(ReturnQty) AS PReturnQty, Sum(costamount) as PCostamount From Ecounts1516.dbo.SaleReport  where TransDate >='2016/04/01' and TransDate <='2016/04/30'"

'ls_sql = ls_sql & "
'ls_sql = "SELECT itemcode, SUM(Amount) - SUM(DiscAmount) AS Saleamount, SUM(SaleQty) AS SaleQty, SUM(SaleReturn) AS SaleReturn, SUM(ReturnQty) AS ReturnQty,SUM(costamount) as Costamount,0  into Tmp_SaleReport From SaleReport"
'ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'  "

'ls_sql = ls_sql & " and convert(varchar,transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
'ls_sql = ls_sql & " and convert(varchar,transdate,111) <= '" & Format(DTPTo.Value, "YYYY/MM/DD") & "' "



'ls_sql = ls_sql & " Group by itemcode "
          
         
'gc_dbcon.Execute ls_sql


'  With crrpt
'          .SQLQuery = ""
'          .Formulas(0) = ""
'          .Formulas(1) = ""
'          .Formulas(2) = ""
'          .Formulas(3) = ""
'          .ReportFileName = App.Path & Gs_ICRepoPath & "\AmjidSbGp.rpt.rpt"
'
'          .WindowTitle = Me.Caption
'          .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
'          .Formulas(1) = "Reportname = 'Item Department Wise Gross Profit Report'"
'          .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
'          .Connect = "DNS=Censoft;UID=Sa"
'          .Action = 1
'   End With

'MDIForm1.StatusBar1.Panels(7).Text = ""
'Exit Sub
'End If
'******************** End Amjid sb Reports ****************



'******************** Start Waste Reports ****************

If txtreporttye.ListIndex = 9 Then


Call ChkTempTables("Tmp_SaleReport", True)

ls_sql = "SELECT itemcode, SUM(Amount) - SUM(DiscAmount) AS Saleamount, SUM(SaleQty) AS SaleQty, SUM(SaleReturn) AS SaleReturn, SUM(ReturnQty) AS ReturnQty,SUM(costamount) as Costamount,SUM(WasteQty) as WasteQty, SUM(WasteAmt) as WasteAmt,sum(LNPCostamount) as LNPCostamount  into Tmp_SaleReport From WasteReport"
ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'  "

ls_sql = ls_sql & " and convert(varchar,transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
ls_sql = ls_sql & " and convert(varchar,transdate,111) <= '" & Format(dtpto.Value, "YYYY/MM/DD") & "' "

If Chktime.Value = 1 Then
ls_sql = ls_sql & " and convert(varchar,transdate1,108) >='" & Format(DTPtimefrom, "HH:mm:ss") & "' "
ls_sql = ls_sql & " and convert(varchar,transdate1,108) <='" & Format(DTPtimeto, "HH:mm:ss") & "' "
End If


If txtcasher.Text <> "" Then
    ls_sql = ls_sql & " and usercode = " & txtcasher.Text & ""
End If
If Chkcash.Value = 1 And chkcredit.Value = 1 Then
    ls_sql = ls_sql & " and Salestatus in(0,1)"
ElseIf Chkcash.Value = 1 Then
    ls_sql = ls_sql & " and Salestatus = 0"
ElseIf chkcredit.Value = 1 Then
    ls_sql = ls_sql & " and Salestatus = 1"
End If
ls_sql = ls_sql & " Group by itemcode "
          
         
gc_dbcon.Execute ls_sql


  With crrpt
          .SQLQuery = ""
          .Formulas(0) = ""
          .Formulas(1) = ""
          .Formulas(2) = ""
          .Formulas(3) = ""
          
          If txtreporttye.ListIndex = 4 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportDeptGDproft.rpt"
          Else
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportDeptGproftWWaste.rpt"
          
          End If
          
          .WindowTitle = Me.Caption
          .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
          .Formulas(1) = "Reportname = 'Item Department Wise Gross Profit Report'"
          .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
          .Connect = "DNS=Censoft;UID=Sa"
          .Action = 1
   End With

MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub
  
 End If
          
'******************** End Waste Reports ****************

MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
If txtreporttye.ListIndex = 3 Or txtreporttye.ListIndex = 4 Or txtreporttye.ListIndex = 10 Then

Call ChkTempTables("Tmp_SaleReport", True)

ls_sql = "SELECT itemcode, SUM(Amount) - SUM(DiscAmount) AS Saleamount, SUM(SaleQty) AS SaleQty, SUM(SaleReturn) AS SaleReturn, SUM(ReturnQty) AS ReturnQty,SUM(costamount) as Costamount,sum(LNPCostamount) as LNPCostamount into Tmp_SaleReport From SaleReport"
ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'  "

ls_sql = ls_sql & " and convert(varchar,transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
ls_sql = ls_sql & " and convert(varchar,transdate,111) <= '" & Format(dtpto.Value, "YYYY/MM/DD") & "' "

If Chktime.Value = 1 Then
ls_sql = ls_sql & " and convert(varchar,transdate1,108) >='" & Format(DTPtimefrom, "HH:mm:ss") & "' "
ls_sql = ls_sql & " and convert(varchar,transdate1,108) <='" & Format(DTPtimeto, "HH:mm:ss") & "' "
End If


If txtcasher.Text <> "" Then
    ls_sql = ls_sql & " and usercode = " & txtcasher.Text & ""
End If
If Chkcash.Value = 1 And chkcredit.Value = 1 Then
    ls_sql = ls_sql & " and Salestatus in(0,1)"
ElseIf Chkcash.Value = 1 Then
    ls_sql = ls_sql & " and Salestatus = 0"
ElseIf chkcredit.Value = 1 Then
    ls_sql = ls_sql & " and Salestatus = 1"
End If
ls_sql = ls_sql & " Group by itemcode "
          
         
gc_dbcon.Execute ls_sql


  With crrpt
          .SQLQuery = ""
          .Formulas(0) = ""
          .Formulas(1) = ""
          .Formulas(2) = ""
          .Formulas(3) = ""
          
          If txtreporttye.ListIndex = 4 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportDeptGDproft.rpt"
          ElseIf txtreporttye.ListIndex = 10 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportDeptGproftSPRate.rpt"
          Else
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportDeptGproft.rpt"
          
          End If
          
          .WindowTitle = Me.Caption
          .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
          .Formulas(1) = "Reportname = 'Item Department Wise Gross Profit Report'"
          .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
          .Connect = "DNS=Censoft;UID=Sa"
          .Action = 1
   End With

MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub
End If


  With crrpt
       .SQLQuery = ""
        If ChkSummary.Value = 1 Then
          If txtreporttye.ListIndex = 0 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportsummarydept.RPT"
          ElseIf txtreporttye.ListIndex = 2 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportsummarycasher.RPT"
          ElseIf txtreporttye.ListIndex = 5 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportsummaryDeptHour.RPT"
          ElseIf txtreporttye.ListIndex = 6 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportsummaryHour.RPT"
          ElseIf txtreporttye.ListIndex = 7 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportDiscsummarydept.RPT"
          ElseIf txtreporttye.ListIndex = 8 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportsummaryusers.RPT"
          End If
          .WindowTitle = Me.Caption
          .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
          .Formulas(1) = "Reportname = 'Sale Report'"
          .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
          
          .SQLQuery = "SELECT SaleReport.Amount, SaleReport.DiscAmount, SaleReport.SaleQty, SaleReport.SaleReturn, SaleReport.ReturnQty, SaleReport.usercode,"
          .SQLQuery = .SQLQuery & " IC_Item.CatCode , SyUsers.UserName, IC_ItemCategory.Description FROM  SaleReport SaleReport LEFT OUTER JOIN"
          .SQLQuery = .SQLQuery & " SyUsers SyUsers ON SaleReport.Compcode = SyUsers.CompCode AND SaleReport.usercode = SyUsers.UserCode LEFT OUTER JOIN"
          .SQLQuery = .SQLQuery & " IC_Item IC_Item ON SaleReport.Compcode = IC_Item.Compcode AND SaleReport.itemcode = IC_Item.ItemCode LEFT OUTER JOIN"
          .SQLQuery = .SQLQuery & " IC_ItemCategory IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode"
          .SQLQuery = .SQLQuery & " where SaleReport.compcode = '" & Gs_compcode & "'  "
          .SQLQuery = .SQLQuery & " and convert(varchar,SaleReport.transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
          .SQLQuery = .SQLQuery & " and convert(varchar,SaleReport.transdate,111) <= '" & Format(dtpto.Value, "YYYY/MM/DD") & "' "
          
           If Chktime.Value = 1 Then
            .SQLQuery = .SQLQuery & " and convert(varchar,SaleReport.transdate1,108) >='" & Format(DTPtimefrom, "HH:mm:ss") & "' "
            .SQLQuery = .SQLQuery & " and convert(varchar,SaleReport.transdate1,108) <='" & Format(DTPtimeto, "HH:mm:ss") & "' "
           End If
        
          
         If txtcasher.Text <> "" Then
         .SQLQuery = .SQLQuery & " and SaleReport.usercode = " & txtcasher.Text & ""
         End If
         If Chkcash.Value = 1 And chkcredit.Value = 1 Then
          .SQLQuery = .SQLQuery & " and SaleReport.Salestatus in(0,1)"
         ElseIf Chkcash.Value = 1 Then
         .SQLQuery = .SQLQuery & " and SaleReport.Salestatus = 0"
         ElseIf chkcredit.Value = 1 Then
         .SQLQuery = .SQLQuery & " and SaleReport.Salestatus = 1"
         End If
         
         If txtreporttye.ListIndex = 7 Then
         .SQLQuery = .SQLQuery & " and SaleReport.discamount  > 0"
           .Formulas(1) = "Reportname = 'Item Wise Sale Discount Report'"
         End If
         If txtvalueadd <> "" Then
           .SQLQuery = .SQLQuery & " " & txtvalueadd
         End If
         .Connect = "DNS=Censoft;UID=Sa"
         .Action = 1
          
        Else
        
        If Chkwinvoicetotal.Value = 1 Then
           .ReportFileName = App.Path & Gs_ICRepoPath & "\SalereportCasher.RPT"
        Else
         .ReportFileName = App.Path & Gs_ICRepoPath & "\SalereportCasherwithoutinv.RPT"
        End If
        .WindowTitle = Me.Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Sale Report'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        .Formulas(3) = "Groupon = " & txtreporttye.ListIndex & ""
 
'        .Formulas(3) = "Groupon = " & txtgroupon.ListIndex + 1 & ""
        .SQLQuery = "SELECT IC_TransMaster.TransCode, IC_TransMaster.TransDate, IC_TransMaster.DiscAmount,IC_TransMaster.SaleStatus,IC_Supplier.Description,"
        .SQLQuery = .SQLQuery & " IC_Trans.Quantity, IC_Trans.Amount, IC_Item.Description, IC_Item.catcode, IC_Item.AvgRate, IC_ItemUM.Description as UOMDEsc,  IC_ItemCategory.Description as Catdesc"
        .SQLQuery = .SQLQuery & " FROM SO_TransMaster IC_TransMaster LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " SO_Trans IC_Trans ON IC_TransMaster.Compcode = IC_Trans.Compcode AND IC_TransMaster.TransCode = IC_Trans.TransCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Clients IC_Supplier ON IC_TransMaster.Compcode = IC_Supplier.Compcode AND"
        .SQLQuery = .SQLQuery & " IC_TransMaster.AccountCode = IC_Supplier.ClientCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Item IC_Item ON IC_Trans.Compcode = IC_Item.Compcode AND IC_Trans.ItemCode = IC_Item.itemcode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_ItemCategory IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.Catcode = IC_ItemCategory.CatCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_ItemUM IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
        .SQLQuery = .SQLQuery & " where IC_TransMaster.Compcode = '" & Gs_compcode & "'"
        .SQLQuery = .SQLQuery & " and convert(varchar,IC_TransMaster.transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
        .SQLQuery = .SQLQuery & " and convert(varchar,IC_TransMaster.transdate,111) <= '" & Format(dtpto.Value, "YYYY/MM/DD") & "' "
        
        If Chktime.Value = 1 Then
            .SQLQuery = .SQLQuery & " and convert(varchar,Ic_TransMaster.transdate,108) >='" & Format(DTPtimefrom, "HH:mm:ss") & "' "
            .SQLQuery = .SQLQuery & " and convert(varchar,Ic_TransMaster.transdate,108) <='" & Format(DTPtimeto, "HH:mm:ss") & "' "
        End If
        
         If Chkcash.Value = 1 And chkcredit.Value = 1 Then
          .SQLQuery = .SQLQuery & " and IC_TransMaster.Salestatus in(0,1)"
         ElseIf Chkcash.Value = 1 Then
         .SQLQuery = .SQLQuery & " and IC_TransMaster.Salestatus = 0"
         ElseIf chkcredit.Value = 1 Then
         .SQLQuery = .SQLQuery & " and IC_TransMaster.Salestatus = 1"
         End If
        
        
         If txtcasher.Text <> "" Then
         .SQLQuery = .SQLQuery & " and IC_TransMaster.usercode = " & txtcasher.Text & ""
         End If
         If txtvalueadd <> "" Then
           .SQLQuery = .SQLQuery & " " & txtvalueadd
         End If
            .SQLQuery = .SQLQuery & " ORDER BY IC_Item.catcode, IC_TransMaster.TransCode"
         .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End If

    End With
   
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command1_Click()
txtvalueadd = txtvalueadd & " " & txtandor & " " & txtfields & " " & txtEq & " " & txtvalue
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 If Chktime.Value = 1 Then
    DTPtimefrom.SetFocus
 Else
    txtcasher1.SetFocus
 End If
End If
End Sub
Private Sub Form_Load()
  dtpfrom = Date
  dtpto = Date
  LoadCasher
  DTPtimefrom = Time
  DTPtimeto = Time
  txtreporttye = "Department Wise"
  Loadcustomreport
End Sub
Private Sub LoadCasher()
Dim pr_loadcasher As New Recordset
pr_loadcasher.Open "SELECT ltrim(rtrim(UserCode)) as UserCode, ltrim(rtrim(UserName)) as UserName  from SyUsers where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loadcasher.EOF Then
Do While Not pr_loadcasher.EOF
   txtcasher1.AddItem pr_loadcasher("Username")
   txtcasher.AddItem pr_loadcasher("UserCode")
pr_loadcasher.MoveNext
Loop
End If
pr_loadcasher.Close
End Sub

Private Sub txtcasher1_Click()
txtcasher.ListIndex = txtcasher1.ListIndex
End Sub

Private Sub txtcasher1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub


