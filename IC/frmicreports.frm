VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmicreport 
   Caption         =   "Sale Register Report"
   ClientHeight    =   6405
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
   Icon            =   "frmicreports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Chktime 
      Caption         =   "Sale With Time"
      Height          =   210
      Left            =   180
      TabIndex        =   15
      Top             =   825
      Width           =   1650
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6030
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
      Height          =   6075
      Left            =   30
      TabIndex        =   5
      Top             =   -90
      Width           =   7980
      Begin VB.Frame Frame4 
         BackColor       =   &H80000016&
         Height          =   735
         Left            =   0
         TabIndex        =   39
         Top             =   2520
         Width           =   7935
         Begin VB.TextBox txtClientdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   240
            Width           =   5250
         End
         Begin VB.CommandButton Command7 
            Height          =   315
            Left            =   2160
            Picture         =   "frmicreports.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtClientCode 
            Height          =   315
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   41
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Client Info Code :"
            Height          =   210
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Update"
         Height          =   330
         Left            =   3315
         TabIndex        =   34
         Top             =   5625
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   6795
         TabIndex        =   33
         Top             =   3360
         Width           =   990
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save "
         Height          =   330
         Left            =   4365
         TabIndex        =   30
         Top             =   5625
         Width           =   1035
      End
      Begin VB.ComboBox txtcustomReport 
         Height          =   330
         ItemData        =   "frmicreports.frx":047C
         Left            =   1545
         List            =   "frmicreports.frx":047E
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3390
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   300
         Left            =   6795
         TabIndex        =   27
         Top             =   3825
         Width           =   990
      End
      Begin VB.TextBox txtvalue 
         Height          =   315
         Left            =   4530
         TabIndex        =   26
         Top             =   3825
         Width           =   2250
      End
      Begin VB.ComboBox txtandor 
         Height          =   330
         ItemData        =   "frmicreports.frx":0480
         Left            =   60
         List            =   "frmicreports.frx":048A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3795
         Width           =   870
      End
      Begin VB.ComboBox txtEq 
         Height          =   330
         ItemData        =   "frmicreports.frx":0497
         Left            =   3525
         List            =   "frmicreports.frx":04A7
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3810
         Width           =   960
      End
      Begin VB.TextBox txtvalueadd 
         Height          =   1200
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   4350
         Width           =   7725
      End
      Begin VB.ComboBox txtfields 
         Height          =   330
         ItemData        =   "frmicreports.frx":04BA
         Left            =   975
         List            =   "frmicreports.frx":04D6
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3810
         Width           =   2565
      End
      Begin VB.CheckBox chksummary 
         Caption         =   "Summary"
         Height          =   300
         Left            =   1770
         TabIndex        =   21
         Top             =   5670
         Width           =   1335
      End
      Begin VB.CheckBox chkwithinvoice 
         Caption         =   "With Invoice Total"
         Height          =   300
         Left            =   90
         TabIndex        =   20
         Top             =   5655
         Width           =   1620
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   6780
         TabIndex        =   4
         Top             =   5625
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
         Left            =   5700
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   5625
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   7440
         Top             =   120
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
         Left            =   4050
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   285
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   2685
         Left            =   0
         TabIndex        =   6
         Top             =   30
         Width           =   7935
         Begin VB.CheckBox ChkLNRate 
            Caption         =   "Report on Last Avg Rate "
            Height          =   495
            Left            =   4680
            TabIndex        =   44
            Top             =   360
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.TextBox txtitemcode 
            Height          =   315
            Left            =   4995
            TabIndex        =   37
            Top             =   2160
            Width           =   720
         End
         Begin VB.TextBox txtitemdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   6015
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1800
         End
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   5700
            Picture         =   "frmicreports.frx":058E
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   2160
            Width           =   315
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   2190
            Picture         =   "frmicreports.frx":0700
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2175
            Width           =   315
         End
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2535
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   2175
            Width           =   1530
         End
         Begin VB.TextBox txtselectedcode 
            Height          =   315
            Left            =   1545
            MaxLength       =   3
            TabIndex        =   16
            Top             =   2175
            Width           =   615
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H80000016&
            Enabled         =   0   'False
            Height          =   1260
            Left            =   30
            TabIndex        =   10
            Top             =   855
            Width           =   7845
            Begin MSComCtl2.DTPicker DTPtimefrom 
               Height          =   315
               Left            =   1545
               TabIndex        =   11
               Top             =   345
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
               Format          =   100728835
               CurrentDate     =   37293
            End
            Begin MSComCtl2.DTPicker DTPtimeto 
               Height          =   315
               Left            =   1545
               TabIndex        =   12
               Top             =   720
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
               Format          =   100728835
               CurrentDate     =   37293
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "From Time :"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   675
               TabIndex        =   14
               Top             =   360
               Width           =   825
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "To Time :"
               Height          =   210
               Left            =   675
               TabIndex        =   13
               Top             =   735
               Width           =   825
            End
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1545
            TabIndex        =   1
            Top             =   180
            Width           =   2475
            _ExtentX        =   4366
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
            Format          =   82837505
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1545
            TabIndex        =   2
            Top             =   540
            Width           =   2490
            _ExtentX        =   4392
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
            Format          =   82837505
            CurrentDate     =   37293
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Item Code :"
            Height          =   210
            Left            =   4170
            TabIndex        =   38
            Top             =   2205
            Width           =   795
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Department Code :"
            Height          =   210
            Left            =   120
            TabIndex        =   19
            Top             =   2205
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   675
            TabIndex        =   8
            Top             =   570
            Width           =   825
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   675
            TabIndex        =   7
            Top             =   195
            Width           =   825
         End
      End
      Begin VB.ComboBox txtcustomReport1 
         Height          =   330
         ItemData        =   "frmicreports.frx":0872
         Left            =   1545
         List            =   "frmicreports.frx":0874
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3390
         Width           =   5250
      End
      Begin VB.ComboBox txtid 
         Height          =   330
         ItemData        =   "frmicreports.frx":0876
         Left            =   1545
         List            =   "frmicreports.frx":0878
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   3390
         Width           =   5280
      End
      Begin VB.Label Label4 
         Caption         =   "Custom Reports :"
         Height          =   300
         Left            =   150
         TabIndex        =   29
         Top             =   3435
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmicreport"
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

Private Sub Check1_Click()

End Sub

Private Sub ChkSummary_Click()
If chksummary.Value = 1 Then
chkwithinvoice.Value = 0
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

Private Sub chkwithinvoice_Click()
If chkwithinvoice.Value = 1 Then
    chksummary.Value = 0
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
'On Error GoTo LocalErr
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
    

   With crrpt
        
        If ChkLNRate.Value = 1 Then
          
          If chkwithinvoice.Value = 1 Then
             .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportall.RPT"
          ElseIf chksummary.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportsummary1.RPT"
          Else
             .ReportFileName = App.Path & Gs_ICRepoPath & "\Salereportwithoutinvoice1.RPT"
          End If
          
        ElseIf chkwithinvoice.Value = 1 Then
           .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportall.RPT"
        ElseIf chksummary.Value = 1 Then
           .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReportsummary.RPT"
        Else
          .ReportFileName = App.Path & Gs_ICRepoPath & "\Salereportwithoutinvoice.RPT"
        End If
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Sale Register'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        
        .SQLQuery = "SELECT IC_TransMaster.TransCode, IC_TransMaster.TransDate, IC_TransMaster.DiscAmount,  IC_Clients.Description,IC_TransMaster.AccounCode, "
        .SQLQuery = .SQLQuery & " IC_Trans.Quantity, IC_Trans.Amount,  IC_Item.Description, IC_Item.Catcode, IC_Item.AvgRate,"
        .SQLQuery = .SQLQuery & " IC_ItemUM.Description,   IC_ItemCategory.Description From  { oj ((((SO_TransMaster IC_TransMaster LEFT OUTER JOIN SO_Trans IC_Trans ON IC_TransMaster.Compcode = IC_Trans.Compcode AND IC_TransMaster.TransCode = IC_Trans.TransCode) LEFT OUTER JOIN IC_Clients IC_Clients ON IC_TransMaster.Compcode = IC_Clients.Compcode AND IC_TransMaster.AccountCode = IC_Clients.ClientCode) LEFT OUTER JOIN IC_Item IC_Item ON IC_Trans.Compcode = IC_Item.Compcode AND IC_Trans.ItemCode = IC_Item.ItemCode)"
        .SQLQuery = .SQLQuery & " LEFT OUTER JOIN IC_ItemCategory  IC_ItemCategory ON IC_Item.Compcode =  IC_ItemCategory.Compcode AND IC_Item.Catcode =  IC_ItemCategory.catCode) LEFT OUTER JOIN IC_ItemUM IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode}"

        '.SQLQuery = "SELECT IC_TransMaster.TransCode, IC_TransMaster.TransDate, IC_TransMaster.DiscAmount,  IC_Clients.Description,"
        '.SQLQuery = .SQLQuery & " IC_Trans.Quantity, IC_Trans.Amount,  IC_Item.Description, IC_Item.Catcode, IC_Item.AvgRate,"
        '.SQLQuery = .SQLQuery & " IC_ItemUM.Description,   IC_ItemCategory.Description From  { oj ((((ecounts.dbo.SO_TransMaster IC_TransMaster LEFT OUTER JOIN ecounts.dbo.SO_Trans IC_Trans ON IC_TransMaster.Compcode = IC_Trans.Compcode AND IC_TransMaster.TransCode = IC_Trans.TransCode) LEFT OUTER JOIN ecounts.dbo.IC_Clients IC_Clients ON IC_TransMaster.Compcode = IC_Clients.Compcode AND IC_TransMaster.AccountCode = IC_Clients.ClientCode) LEFT OUTER JOIN ecounts.dbo.IC_Item IC_Item ON IC_Trans.Compcode = IC_Item.Compcode AND IC_Trans.ItemCode = IC_Item.ItemCode)"
        '.SQLQuery = .SQLQuery & " LEFT OUTER JOIN ecounts.dbo. IC_ItemCategory  IC_ItemCategory ON IC_Item.Compcode =  IC_ItemCategory.Compcode AND IC_Item.Catcode =  IC_ItemCategory.catCode) LEFT OUTER JOIN ecounts.dbo.IC_ItemUM IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode}"

        .SQLQuery = .SQLQuery & " where Ic_TransMaster.CompCode ='" & Gs_compcode & "'"
        .SQLQuery = .SQLQuery & " and convert(varchar,Ic_TransMaster.transdate,111) >='" & Format(dtpfrom, "YYYY/MM/DD") & "' "
        .SQLQuery = .SQLQuery & " and convert(varchar,Ic_TransMaster.transdate,111) <='" & Format(dtpto, "YYYY/MM/DD") & "' "
        
        If Chktime.Value = 1 Then
        .SQLQuery = .SQLQuery & " and convert(varchar,Ic_TransMaster.transdate,108) >='" & Format(DTPtimefrom, "HH:mm:ss") & "' "
        .SQLQuery = .SQLQuery & " and convert(varchar,Ic_TransMaster.transdate,108) <='" & Format(DTPtimeto, "HH:mm:ss") & "' "
        End If
        If txtvalueadd <> "" Then
        .SQLQuery = .SQLQuery & " " & txtvalueadd
        End If
        
        If txtClientCode <> "" Then
        .SQLQuery = .SQLQuery & " and  Ic_TransMaster.AccountCode ='" & Trim(txtClientCode) & "'"
        End If
        
        If txtselectedcode <> "" Then
        .SQLQuery = .SQLQuery & " and  IC_ItemCategory.CatCode ='" & txtselectedcode & "'"
        End If
  
        If txtitemcode <> "" Then
        .SQLQuery = .SQLQuery & " and   IC_Trans.itemCode ='" & txtitemcode & "'"
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
'If txtvalueadd = "" Then
'txtvalueadd = txtvalueadd & " " & txtandor & " " & txtvalueadd & " " & txtfields & " " & txtEq & " " & txtvalue
'Else
txtvalueadd = txtvalueadd & " " & txtandor & " " & txtfields & " " & txtEq & " " & txtvalue
'End If
End Sub

Private Sub Command2_Click()
Dim res
res = InputBox("Name of the Report", "Custom Reort")
If res <> "" Then
gc_dbcon.Execute "Insert  into so_customreport (reportname,parameter,reportid) values('" & res & "','" & txtvalueadd & "',1) "
Loadcustomreport
End If

End Sub

Private Sub Command3_Click()
Dim ls_id
ls_id = txtid.Text
gc_dbcon.Execute "delete from so_Customreport where id = " & ls_id & " and reportid = 1 "
Loadcustomreport
End Sub

Private Sub Command4_Click()
On Error GoTo localerror
Dim ls_id
ls_id = txtid.Text
gc_dbcon.Execute "delete from so_Customreport where id = " & ls_id & " and reportid = 1 "
gc_dbcon.Execute "Insert  into so_customreport (reportname,parameter,reportid) values('" & txtcustomReport.Text & "','" & txtvalueadd & "',1) "
Loadcustomreport
Exit Sub
localerror:

End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectedcode
    Set PO_DESC = txtdesc
    Gs_SQL = "SELECT Catcode,Description  from IC_ItemCategory"
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    If txtselectedcode <> "" Then Call txtselectedcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command6_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtitemcode
    Set PO_DESC = txtitemdesc
    Gs_SQL = "Select IC_Item.ItemCode,   IC_Item.Description, IC_ItemCategory.Description as Category,IC_Item.SaleCost from IC_Item left outer join IC_ItemCategory on IC_Item.compcode = IC_ItemCategory.compcode and   IC_Item.catcode = IC_ItemCategory.catcode "
    Gs_FindFld = "IC_Item.Description"
    Gs_OrderBy = "Order by IC_Item.Description"
    Gs_OtherPara = " where IC_Item.compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1

    If Len(txtitemcode) > 0 Then txtItemcode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command7_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtClientCode
    Set PO_DESC = txtClientdesc
    Gs_SQL = "Select ClientCode, Description from IC_Clients "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Supplier"
    MyLookupOLDB.Show 1

    If Len(txtClientCode) > 0 Then txtClientCode_KeyDown vbKeyReturn, vbKeyShift


End Sub

Private Sub txtClientCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Trim(txtClientCode.Text) <> "" Then
       txtClientCode.Text = DoPad(UCase(txtClientCode.Text), txtClientCode.MaxLength)
       If pr_dumy.State = 1 Then pr_dumy.Close
       pr_dumy.Open "Select * from IC_Clients where Compcode  = '" & Gs_compcode & "' and ClientCode = '" & txtClientCode.Text & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
               If Not pr_dumy.EOF Then
                   txtClientCode = pr_dumy("ClientCode")
                   txtClientdesc = pr_dumy("Description")
                   txtClientCode.SetFocus
               
                End If
           
        pr_dumy.Close
  
End If
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub
Private Sub Loadcustomreport()
Dim pr_CustomReport As New Recordset
   txtcustomReport.Clear
   txtcustomReport1.Clear
   txtid.Clear

pr_CustomReport.Open "SELECT *  from So_customreport where reportid = 1", gc_dbcon, adOpenStatic, adLockReadOnly, 1
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

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If Chktime.Value = 1 Then
    DTPtimefrom.SetFocus
Else
    txtselectedcode.SetFocus
End If
End If
End Sub

Private Sub Form_Load()
  dtpfrom = Date
  dtpto = Date
  DTPtimefrom = Time
  DTPtimeto = Time
  Loadcustomreport
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
Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtitemcode <> "" And KeyCode = vbKeyReturn Then
    If txtitemcode <> "Selective" Then
    ls_sql = "Select itemcode,Description from IC_Item where compcode = '" & Gs_compcode & "' and Itemcode= '" & txtitemcode & "' "
    
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Item Code not found", vbCritical)
            Else
                 txtitemdesc = pr_dumy("description")
                 cmdGenerate.SetFocus
            End If
         pr_dumy.Close
    End If
End If

End Sub

Private Sub txtselectedcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtselectedcode <> "" Then
    txtselectedcode = DoPad(txtselectedcode, txtselectedcode.MaxLength)
    ls_sql = "Select Catcode,Description from IC_ItemCategory where compcode = '" & Gs_compcode & "' and Catcode = '" & txtselectedcode & "' "
    pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
     If pr_dumy.EOF Then
            Call MsgBox("Department code not found !!!", vbCritical)
            txtselectedcode = ""
            txtdesc = ""
            txtselectedcode.SetFocus
     Else
                txtdesc = pr_dumy("description")
                cmdGenerate.SetFocus
     End If
          pr_dumy.Close

ElseIf KeyCode = vbKeyReturn And txtselectedcode = "" Then
        Command5_Click
End If

End Sub
