VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSosearchRecords 
   Caption         =   "Look up :"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSoSearchReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptVoucher 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Print Invoice"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12585
      TabIndex        =   10
      Top             =   60
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4785
      TabIndex        =   6
      Top             =   -45
      Width           =   7725
      Begin VB.TextBox SeekText 
         Height          =   315
         Left            =   525
         MaxLength       =   50
         TabIndex        =   7
         Top             =   195
         Width           =   7110
      End
      Begin VB.Label Label2 
         Caption         =   "Text :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   210
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdtree 
      Caption         =   "&Tree"
      Height          =   300
      Left            =   5565
      TabIndex        =   5
      Top             =   5790
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6825
      Left            =   30
      TabIndex        =   3
      Top             =   585
      Width           =   14805
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6570
         Left            =   105
         TabIndex        =   0
         Top             =   195
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   11589
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
         TabAction       =   1
         WrapCellPointer =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   105
      TabIndex        =   1
      Top             =   -45
      Width           =   4650
      Begin VB.ComboBox txtsearchby 
         Height          =   330
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   195
         Width           =   3570
      End
      Begin VB.Label Label1 
         Caption         =   "Search BY :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8205
      Top             =   -105
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblmsg 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   45
      TabIndex        =   4
      Top             =   2895
      Width           =   45
   End
End
Attribute VB_Name = "frmSosearchRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ln_res As Integer
Dim PR_Sub1 As New Recordset
Public ls_Accountno As String


Private Sub Command1_Click()
Printinvoice
End Sub
Private Sub Printinvoice()
On Error GoTo LocalErr
If Me.Caption = "Invoices" Then
   With rptVoucher
        .WindowTitle = Me.Caption
        If ln_changeprinter = 1 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "AuoraSaleInvoice.rpt"
        Else
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SaleInvoice.rpt"
        End If
        
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
         .SQLQuery = "SELECT SO_TransMaster.TransCode, SO_TransMaster.TransDate, SO_TransMaster.DiscAmount, SO_TransMaster.RecAmount, SO_TransMaster.BalAmount, "
        .SQLQuery = .SQLQuery & " SO_TransMaster.CompName , SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount, SyUsers.UserName, IC_Item.Description,IC_Clients.Description"
        .SQLQuery = .SQLQuery & " FROM SO_TransMaster SO_TransMaster LEFT OUTER JOIN SyUsers SyUsers ON SO_TransMaster.Compcode = SyUsers.CompCode AND SO_TransMaster.UserCode = SyUsers.UserCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " SO_Trans SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND SO_TransMaster.TransCode = SO_Trans.TransCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Item IC_Item ON SO_Trans.Compcode = IC_Item.Compcode AND SO_Trans.ItemCode = IC_Item.ItemCode  "
        .SQLQuery = .SQLQuery & " LEFT OUTER JOIN IC_Clients IC_Clients ON SO_TransMaster.Compcode = IC_Clients.Compcode AND SO_TransMaster.AccountCode = IC_Clients.ClientCode"
        .SQLQuery = .SQLQuery & " where SO_TransMaster.compcode = '" & Gs_compcode & "' and  SO_TransMaster.transcode = '" & Trim(DataGrid1.Columns(0).Text) & "'"
        .SQLQuery = .SQLQuery & " ORDER BY SO_TransMaster.TransCode "
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
ElseIf Me.Caption = "SALE RETURN" Then

   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "AuoraSaleInvoicereturn.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        '.Formulas(2) = "Reportname = 'Good Receive Note'"
        .SQLQuery = "SELECT SO_TransMaster.TransCode, SO_TransMaster.TransDate, SO_TransMaster.DiscAmount, SO_TransMaster.RecAmount, SO_TransMaster.BalAmount, "
        .SQLQuery = .SQLQuery & " SO_TransMaster.CompName ,SO_Trans.discamt, SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount,IC_Clients.Description, SyUsers.UserName, IC_Item.Description"
        .SQLQuery = .SQLQuery & " FROM SO_TransReturnMaster SO_TransMaster LEFT OUTER JOIN SyUsers SyUsers ON SO_TransMaster.Compcode = SyUsers.CompCode AND SO_TransMaster.UserCode = SyUsers.UserCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " SO_TransReturn SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND SO_TransMaster.TransCode = SO_Trans.TransCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Item IC_Item ON SO_Trans.Compcode = IC_Item.Compcode AND SO_Trans.ItemCode = IC_Item.ItemCode"
        .SQLQuery = .SQLQuery & " LEFT OUTER JOIN  IC_Clients IC_Clients ON SO_TransMaster.Compcode = IC_Clients.Compcode AND SO_TransMaster.AccountCode = IC_Clients.ClientCode"
        .SQLQuery = .SQLQuery & " where SO_TransMaster.compcode = '" & Gs_compcode & "' and  SO_TransMaster.transcode = '" & Trim(DataGrid1.Columns(0).Text) & "'"
        .SQLQuery = .SQLQuery & " ORDER BY SO_TransMaster.TransCode "
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With




End If


Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub

Private Sub DataGrid1_DblClick()
 Printinvoice
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim sortField As String
Dim sortString As String

sortField = DataGrid1.Columns(ColIndex).Caption
If InStr(Adodc1.Recordset.Sort, "Asc") Then
    sortString = sortField & " Desc"
Else
    sortString = sortField & " Asc"
End If
Adodc1.Recordset.Sort = sortString
End Sub


Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    If Adodc1.Recordset.RecordCount >= 1 Then
        PO_AnyForm.PO_CODE = Trim(DataGrid1.Columns(0).Text)
        PO_AnyForm.PO_DESC = Trim(DataGrid1.Columns(1).Text)
        Unload Me
    End If
ElseIf KeyCode = vbKeyPageUp Then
    SeekText.SetFocus
End If

End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    If Adodc1.Recordset.RecordCount > 1 Then DataGrid1_KeyDown vbKeyReturn, vbKeyShift
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
'On Error GoTo LocalErr

Adodc1.ConnectionString = gc_dbcon
Adodc1.CommandType = adCmdText


Adodc1.RecordSource = Gs_SQL + Gs_OtherPara + " " + Gs_OrderBy
Adodc1.Refresh

txtsearchby.AddItem Adodc1.Recordset.Fields(2).Name
txtsearchby.AddItem Adodc1.Recordset.Fields(3).Name
txtsearchby.AddItem Adodc1.Recordset.Fields(4).Name
txtsearchby.Text = Adodc1.Recordset.Fields(4).Name
Gs_FindFld = Adodc1.Recordset.Fields(4).Name
Set DataGrid1.DataSource = Adodc1

DataGrid1.Columns(2).Width = 4000
DataGrid1.Columns(3).Width = 4000
        
 If Gs_Subon = True Then
    PR_Sub1.Open "select gl_sub2.*,ltrim(rtrim(gl_sub2.acct_sub1))+ltrim(rtrim(gl_sub2.acct_sub2)) as Findfld from gl_sub2 where compcode = '" & Gs_compcode & "' order by  Findfld ", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
    cmdtree.Visible = True
 End If
    Exit Sub
LocalErr:
ln_res = SetErr("Critical error occurred Please report to MIS Department", vbCritical)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Gs_Subon = True Then
         PR_Sub1.Close
    End If
    Gs_SQL = ""
    Gs_Subon = False
    Gs_FindFld = ""
    Gs_OrderBy = ""
    Gs_OtherPara = ""
    
End Sub

Private Sub SeekText_Change()
On Error GoTo LocalErr
If SeekText.Text <> "" Then
   'Gs_OtherPara = " Where Transdate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <= '" & Format(DTPTo, "YYYY/MM/DD") & "'"
   Adodc1.RecordSource = Gs_SQL & IIf(Len(Gs_OtherPara) > 0, Gs_OtherPara & " And " & UCase(Gs_FindFld) & " like '" & "%" & UCase(SeekText) & "%" & "'" + " " + Gs_OrderBy, " Where " & UCase(Gs_FindFld) & " like '" & UCase(SeekText) & "%" & "'" + " " + Gs_OrderBy)
   Adodc1.Refresh
Else
    Adodc1.RecordSource = Gs_SQL + Gs_OtherPara + " " + Gs_OrderBy
    Adodc1.Refresh
End If

DataGrid1.Columns(2).Width = 4000
DataGrid1.Columns(3).Width = 4000
 
Exit Sub
LocalErr:
ln_res = SetErr("Critical error occurred Please report to MIS Department", vbCritical)
End Sub

Private Sub SeekText_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then DataGrid1.SetFocus
End Sub

Private Sub txtsearchby_Click()
Gs_FindFld = txtsearchby.Text
End Sub

Private Sub txtsearchby_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SeekText.SetFocus
End Sub
