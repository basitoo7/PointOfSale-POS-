VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmDashboard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dash Board"
   ClientHeight    =   10950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20250
   Icon            =   "DashBoard.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Caption         =   "Monthly Net Sale Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   13785
      TabIndex        =   11
      Top             =   7230
      Width           =   6360
      Begin MSChart20Lib.MSChart MSChart5 
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "DashBoard.frx":030A
         TabIndex        =   16
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Current Month Net Sale Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   13770
      TabIndex        =   10
      Top             =   30
      Width           =   6360
      Begin MSChart20Lib.MSChart MSChart4 
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "DashBoard.frx":2660
         TabIndex        =   15
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   13920
      Top             =   135
   End
   Begin VB.Frame Frame6 
      Caption         =   "Monthly Sale Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   5310
      TabIndex        =   8
      Top             =   7230
      Width           =   8460
      Begin MSChart20Lib.MSChart MSChart3 
         Height          =   3135
         Left            =   120
         OleObjectBlob   =   "DashBoard.frx":49B6
         TabIndex        =   14
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Daily Category Wise Sale Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   5280
      TabIndex        =   7
      Top             =   3600
      Width           =   14865
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   3135
         Left            =   120
         OleObjectBlob   =   "DashBoard.frx":6D0C
         TabIndex        =   13
         Top             =   360
         Width           =   14415
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Current Month Sale Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   5295
      TabIndex        =   6
      Top             =   30
      Width           =   8460
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "DashBoard.frx":9062
         TabIndex        =   12
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Monthly Sale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   0
      TabIndex        =   4
      Top             =   7215
      Width           =   5205
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN2 
         Height          =   3165
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   5583
         _Version        =   393216
         BackColor       =   16777215
         RowHeightMin    =   300
         BackColorSel    =   16777215
         ForeColorSel    =   0
         GridColor       =   -2147483632
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Daily Category Wise Sale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   5205
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN1 
         Height          =   3165
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   5583
         _Version        =   393216
         BackColor       =   16777215
         RowHeightMin    =   500
         BackColorSel    =   16777215
         ForeColorSel    =   0
         GridColor       =   -2147483632
         WordWrap        =   -1  'True
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current Month Sale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   5205
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3660
         MaskColor       =   &H00CB9E6B&
         TabIndex        =   9
         Top             =   3060
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   2715
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4789
         _Version        =   393216
         BackColor       =   16777215
         RowHeightMin    =   300
         BackColorSel    =   16777215
         ForeColorSel    =   0
         GridColor       =   -2147483632
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Description|<Sale|<Sale Return"
        .ColWidth(1) = 1700
        .ColWidth(2) = 1400
        .ColWidth(3) = 1430
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .Redraw = True
    End With
End Sub

Private Sub InitializeGrid1()
    With GrdGRN1
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Description|<Sale|<Sale Return"
        .ColWidth(1) = 1700
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .Redraw = True
    End With
End Sub
Private Sub InitializeGrid2()
    With GrdGRN2
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Description|<Sale|<Sale Return"
        .ColWidth(1) = 1700
        .ColWidth(2) = 1400
        .ColWidth(3) = 1430
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .Redraw = True
    End With
End Sub
Private Sub LoadcurrentMonthData()
On Error GoTo LocalErr
Dim ls_sql As String
Dim ld_date As Date
Dim ls_Tsaleamount As Double
Dim ls_tsaleretamount As Double

ls_Tsaleamount = 0
ls_tsaleretamount = 0

Dim pr_loaddata As New Recordset
ls_sql = "SELECT convert(varchar(20),transdate,103) as DateDesc, sum(Amount)- sum(DiscAmount) as Amount,sum(SaleReturn) as SaleReturn FROM   SaleReport where  compcode = '" & Gs_compcode & "'  and convert(varchar,transdate,111) >= '" & Format(DateAdd("d", -3, Gd_SysDate), "YYYY/MM/DD") & "'  and convert(varchar,transdate,111) <= '" & Format(Gd_SysDate, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by convert(varchar,transdate,111)"
ls_sql = ls_sql & "  order by   convert(varchar,transdate,111) desc"

If pr_loaddata.State = 1 Then pr_loaddata.Close
pr_loaddata.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loaddata.EOF Then
MSChart1.Visible = True
MSChart4.Visible = True
With GrdGRN
Do While Not pr_loaddata.EOF
    .Row = .Rows - 1
    .TextMatrix(.Row, 1) = Format(pr_loaddata("DateDesc"), "DD/MM/YYYY")
    
    .TextMatrix(.Row, 2) = Format(Val(0 & pr_loaddata("Amount")), "#,##0;($#,##0)")
    .TextMatrix(.Row, 3) = Format(Val(0 & pr_loaddata("SaleReturn")), "#,##0;($#,##0)")
     ls_Tsaleamount = ls_Tsaleamount + Val(0 & pr_loaddata("Amount"))
     ls_tsaleretamount = ls_tsaleretamount + Val(0 & pr_loaddata("SaleReturn"))
    .Rows = .Rows + 1

pr_loaddata.MoveNext
Loop
        If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
End With
Else
MSChart1.Visible = False
MSChart4.Visible = False
End If
pr_loaddata.Close
DoEvents

ls_sql = "SELECT  convert(varchar(20),transdate,103) as DateDesc ,((sum(Amount)- sum(DiscAmount))/" & ls_Tsaleamount & " )*100  as SaleAmount,(sum(SaleReturn)/ (sum(Amount)- sum(DiscAmount)) )* 100 as SaleReturn FROM   SaleReport where compcode = '" & Gs_compcode & "'  and convert(varchar,transdate,111) >= '" & Format(DateAdd("d", -3, Gd_SysDate), "YYYY/MM/DD") & "'  and convert(varchar,transdate,111) <= '" & Format(Gd_SysDate, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by convert(varchar,transdate,111)"
ls_sql = ls_sql & "  order by   convert(varchar,transdate,111) desc"

If pr_loaddata.State = 1 Then pr_loaddata.Close
pr_loaddata.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loaddata.EOF Then
Set MSChart1.DataSource = pr_loaddata

MSChart1.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set 66, 134, 244
MSChart1.Plot.SeriesCollection(2).DataPoints(-1).Brush.FillColor.Set 198, 122, 59


MSChart1.ShowLegend = True
Dim rownumber As Integer
rownumber = 1
Do While Not pr_loaddata.EOF
MSChart1.Row = rownumber
MSChart1.RowLabel = Format(pr_loaddata("DateDesc"), "DD/MM/YYYY")
rownumber = rownumber + 1
pr_loaddata.MoveNext
Loop

MSChart1.Title = "Net Sale And Return For Current Month"

End If
pr_loaddata.Close
DoEvents

'Net Sale

Frame7.Caption = "Loading..."

DoEvents

ls_sql = "SELECT  convert(varchar(20),transdate,103) as DateDesc ,((sum(Amount)- sum(DiscAmount))/" & (ls_Tsaleamount - ls_tsaleretamount) & " )*100  as SaleAmount FROM   SaleReport where compcode = '" & Gs_compcode & "'  and convert(varchar,transdate,111) >= '" & Format(DateAdd("d", -3, Gd_SysDate), "YYYY/MM/DD") & "'  and convert(varchar,transdate,111) <= '" & Format(Gd_SysDate, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by convert(varchar,transdate,111)"
ls_sql = ls_sql & "  order by   convert(varchar,transdate,111) desc"

If pr_loaddata.State = 1 Then pr_loaddata.Close
pr_loaddata.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loaddata.EOF Then
Set MSChart4.DataSource = pr_loaddata

MSChart4.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set 196, 125, 158





rownumber = 1
Do While Not pr_loaddata.EOF
MSChart4.Row = rownumber
MSChart4.RowLabel = Format(pr_loaddata("DateDesc"), "DD/MM/YYYY")
rownumber = rownumber + 1
pr_loaddata.MoveNext
Loop

MSChart4.Title = "Net Sale For Current Month"

End If
pr_loaddata.Close
DoEvents

Frame7.Caption = "Current Month Net Sale Graph"

Exit Sub
LocalErr:

Call MsgBox(Err.Description)

End Sub

Private Sub LoadthreeMonthData()
On Error GoTo LocalErr
Dim ls_sql As String
Dim ld_date As Date
Dim ls_Tsaleamount As Double
Dim ls_tsaleretamount As Double

ls_Tsaleamount = 0
ls_tsaleretamount = 0

Dim pr_loaddata As New Recordset
ls_sql = "SELECT RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   RIGHT('00' + CAST(MONTH(transdate) as VARCHAR),2) as Monthnum,RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   LEFT(DATENAME(MONTH,transdate),3) as DateDesc, sum(Amount)- sum(DiscAmount) as Amount,sum(SaleReturn) as SaleReturn FROM   SaleReport where compcode = '" & Gs_compcode & "'  and convert(varchar,transdate,111) >= '" & Format(DateAdd("m", -3, Gd_SysDate), "YYYY/MM/DD") & "'  and convert(varchar,transdate,111) <= '" & Format(Gd_SysDate, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   RIGHT('00' + CAST(MONTH(transdate) as VARCHAR),2) ,RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   LEFT(DATENAME(MONTH,transdate),3)"
ls_sql = ls_sql & "  order by   RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   RIGHT('00' + CAST(MONTH(transdate) as VARCHAR),2) desc"

If pr_loaddata.State = 1 Then pr_loaddata.Close
pr_loaddata.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loaddata.EOF Then
MSChart3.Visible = True
MSChart5.Visible = True

With GrdGRN2
Do While Not pr_loaddata.EOF
    .Row = .Rows - 1
    .TextMatrix(.Row, 1) = pr_loaddata("DateDesc")
    
    .TextMatrix(.Row, 2) = Format(Val(0 & pr_loaddata("Amount")), "#,##0;($#,##0)")
    .TextMatrix(.Row, 3) = Format(Val(0 & pr_loaddata("SaleReturn")), "#,##0;($#,##0)")
     ls_Tsaleamount = ls_Tsaleamount + Val(0 & pr_loaddata("Amount"))
     ls_tsaleretamount = ls_tsaleretamount + Val(0 & pr_loaddata("SaleReturn"))
    .Rows = .Rows + 1

pr_loaddata.MoveNext
Loop
        If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
End With
Else
MSChart3.Visible = False
MSChart5.Visible = False

End If
pr_loaddata.Close
DoEvents



ls_sql = "SELECT  RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   LEFT(DATENAME(MONTH,transdate),3) as DateDesc ,((sum(Amount)- sum(DiscAmount))/" & ls_Tsaleamount & " )*100  as SaleAmount,(sum(SaleReturn)/ (sum(Amount)- sum(DiscAmount)) )* 100 as SaleReturn FROM   SaleReport where compcode = '" & Gs_compcode & "'  and convert(varchar,transdate,111) >= '" & Format(DateAdd("m", -3, Gd_SysDate), "YYYY/MM/DD") & "'  and convert(varchar,transdate,111) <= '" & Format(Gd_SysDate, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   RIGHT('00' + CAST(MONTH(transdate) as VARCHAR),2) , RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   LEFT(DATENAME(MONTH,transdate),3)"
ls_sql = ls_sql & "  order by  RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   RIGHT('00' + CAST(MONTH(transdate) as VARCHAR),2) desc"

If pr_loaddata.State = 1 Then pr_loaddata.Close
pr_loaddata.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loaddata.EOF Then
Set MSChart3.DataSource = pr_loaddata

MSChart3.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set 66, 134, 244
MSChart3.Plot.SeriesCollection(2).DataPoints(-1).Brush.FillColor.Set 198, 122, 59


MSChart3.ShowLegend = True
Dim rownumber As Integer
rownumber = 1
Do While Not pr_loaddata.EOF
MSChart3.Row = rownumber
MSChart3.RowLabel = pr_loaddata("DateDesc")
rownumber = rownumber + 1
pr_loaddata.MoveNext
Loop



End If
MSChart3.Title = "Net Sale And Return For Three Months"

pr_loaddata.Close
DoEvents


'net sale loading
Frame8.Caption = "Loading..."
DoEvents
ls_sql = "SELECT  RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   LEFT(DATENAME(MONTH,transdate),3) as DateDesc ,((sum(Amount)- sum(DiscAmount))/" & (ls_Tsaleamount - ls_tsaleretamount) & " )*100  as NetSaleAmount FROM   SaleReport where compcode = '" & Gs_compcode & "'  and convert(varchar,transdate,111) >= '" & Format(DateAdd("m", -3, Gd_SysDate), "YYYY/MM/DD") & "'  and convert(varchar,transdate,111) <= '" & Format(Gd_SysDate, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   RIGHT('00' + CAST(MONTH(transdate) as VARCHAR),2) , RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   LEFT(DATENAME(MONTH,transdate),3)"
ls_sql = ls_sql & "  order by  RIGHT('00' + CAST(YEAR(transdate) AS VARCHAR),2 )+ '-' +   RIGHT('00' + CAST(MONTH(transdate) as VARCHAR),2) desc"

If pr_loaddata.State = 1 Then pr_loaddata.Close
pr_loaddata.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loaddata.EOF Then
Set MSChart5.DataSource = pr_loaddata

MSChart5.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set 196, 125, 158


rownumber = 1
Do While Not pr_loaddata.EOF
MSChart5.Row = rownumber
MSChart5.RowLabel = pr_loaddata("DateDesc")
rownumber = rownumber + 1
pr_loaddata.MoveNext
Loop


End If
pr_loaddata.Close
DoEvents
MSChart5.Title = "Net Sale For Three Months"
Frame8.Caption = "Net Sale For Three Months"
Exit Sub
LocalErr:

Call MsgBox(Err.Description)
End Sub

Private Sub LoadcurrentMonthcatData()
On Error GoTo LocalErr
Dim ls_sql As String
Dim ld_date As Date
Dim ls_Tsaleamount As Double
Dim ls_tsaleretamount As Double

ls_Tsaleamount = 0
ls_tsaleretamount = 0

Dim pr_loaddata As New Recordset

ls_sql = "SELECT isnull(c.description,'Desc Please') as DateDesc, sum(a.Amount)- sum(a.DiscAmount) as Amount,sum(a.SaleReturn) as SaleReturn FROM   SaleReport a"
ls_sql = ls_sql & " left outer join ic_item b on a.itemcode = b.itemcode and a.compcode = b.compcode"
ls_sql = ls_sql & " left outer join IC_ItemCategory c  on b.catcode = c.catcode and b.compcode =c.compcode"
ls_sql = ls_sql & " where a.compcode = '" & Gs_compcode & "'   and convert(varchar,a.transdate,111) >= '" & Format(Gd_SysDate, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by  c.description order by c.description "

If pr_loaddata.State = 1 Then pr_loaddata.Close
pr_loaddata.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loaddata.EOF Then
MSChart2.Visible = True


With GrdGRN1
Do While Not pr_loaddata.EOF
    .Row = .Rows - 1
    .TextMatrix(.Row, 1) = pr_loaddata("DateDesc")
    
    .TextMatrix(.Row, 2) = Format(Val(0 & pr_loaddata("Amount")), "#,##0;($#,##0)")
    .TextMatrix(.Row, 3) = Format(Val(0 & pr_loaddata("SaleReturn")), "#,##0;($#,##0)")
     ls_Tsaleamount = ls_Tsaleamount + Val(0 & pr_loaddata("Amount"))
     ls_tsaleretamount = ls_tsaleretamount + Val(0 & pr_loaddata("SaleReturn"))
    .Rows = .Rows + 1

pr_loaddata.MoveNext
Loop
        If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
End With
Else
MSChart2.Visible = False


End If
pr_loaddata.Close
DoEvents

ls_Tsaleamount = ls_Tsaleamount - ls_tsaleretamount

ls_sql = "SELECT isnull(c.description,'Desc Please') as DateDesc, ((sum(Amount)- (sum(DiscAmount)+sum(SaleReturn)))/" & ls_Tsaleamount & " )*100  as NetSaleAmount   FROM   SaleReport a"
ls_sql = ls_sql & " left outer join ic_item b on a.itemcode = b.itemcode and a.compcode = b.compcode"
ls_sql = ls_sql & " left outer join IC_ItemCategory c  on b.catcode = c.catcode and b.compcode =c.compcode"
ls_sql = ls_sql & " where a.compcode = '" & Gs_compcode & "'   and convert(varchar,a.transdate,111) >= '" & Format(Gd_SysDate, "YYYY/MM/DD") & "'"
ls_sql = ls_sql & " group by  c.description order by c.description "



If pr_loaddata.State = 1 Then pr_loaddata.Close
pr_loaddata.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loaddata.EOF Then
Set MSChart2.DataSource = pr_loaddata

MSChart2.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Set 66, 134, 244
'MSChart2.Plot.SeriesCollection(2).DataPoints(-1).Brush.FillColor.Set 198, 122, 59


MSChart2.ShowLegend = True
Dim rownumber As Integer
rownumber = 1
Do While Not pr_loaddata.EOF
MSChart2.Row = rownumber
MSChart2.RowLabel = Left(pr_loaddata("DateDesc"), 5)
rownumber = rownumber + 1
pr_loaddata.MoveNext
Loop


End If

MSChart2.Title = "Net Sale Daily Category Wise"

pr_loaddata.Close
DoEvents

Exit Sub
LocalErr:

Call MsgBox(Err.Description)
End Sub

Private Sub Command1_Click()
Timer1_Timer
End Sub

Private Sub Timer1_Timer()
InitializeGrid
InitializeGrid2
InitializeGrid1
DoEvents


Frame2.Caption = "Loading..."
Frame4.Caption = "Loading..."
DoEvents
LoadcurrentMonthData
Frame4.Caption = "Current Month Sale"
Frame2.Caption = "Current Month Sale Graph"
DoEvents

Frame1.Caption = "Loading..."
Frame5.Caption = "Loading..."
LoadcurrentMonthcatData
Frame5.Caption = "Daily Category Wise Sale Graph"
Frame1.Caption = "Daily Category Wise Sale"

DoEvents

Frame3.Caption = "Loading..."
Frame6.Caption = "Loading..."
LoadthreeMonthData
Frame6.Caption = "Monthly Sale Graph"
Frame3.Caption = "Monthly Sale"

DoEvents
Timer1.Interval = 0
End Sub
