VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOVendorPaymentsBank 
   Caption         =   "Payments to Vendors By Bank "
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendorPaymentsBank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
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
      Height          =   6570
      Left            =   0
      TabIndex        =   2
      Top             =   555
      Width           =   11820
      Begin VB.TextBox txttotalamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   5985
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6150
         Width           =   1470
      End
      Begin VB.TextBox txtNetamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   10290
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   6150
         Width           =   1470
      End
      Begin VB.TextBox txtbalanceAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   6165
         Width           =   1470
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   2070
         Picture         =   "frmVendorPaymentsBank.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   510
         Width           =   315
      End
      Begin VB.TextBox txtbankcode 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1395
         MaxLength       =   4
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   660
      End
      Begin VB.TextBox txtbankdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   2415
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   4650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Load GRN"
         Height          =   300
         Left            =   5595
         TabIndex        =   13
         Top             =   150
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtremarks 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1395
         MaxLength       =   200
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Text            =   "CHQ. ISSUED TO SUPPLIERS"
         Top             =   900
         Width           =   10305
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   11565
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   195
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2505
         Picture         =   "frmVendorPaymentsBank.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   135
         Width           =   315
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
         Height          =   330
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   -15
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   195
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   3915
         TabIndex        =   0
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63045633
         CurrentDate     =   37580
      End
      Begin Crystal.CrystalReport rptVoucher 
         Left            =   45
         Top             =   1875
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         Destination     =   1
         WindowBorderStyle=   3
         WindowControlBox=   0   'False
         WindowMaxButton =   0   'False
         WindowMinButton =   0   'False
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Frame Frame2 
         Height          =   4950
         Left            =   75
         TabIndex        =   8
         Top             =   1170
         Width           =   11700
         Begin VB.TextBox TXTBARCODE 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            MaxLength       =   255
            TabIndex        =   18
            Tag             =   "SKIP"
            Top             =   0
            Visible         =   0   'False
            Width           =   2310
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
            Height          =   4620
            Left            =   75
            TabIndex        =   9
            Top             =   195
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   8149
            _Version        =   393216
            RowHeightMin    =   300
            BackColorSel    =   16777215
            ForeColorSel    =   0
            GridColor       =   8421504
            AllowBigSelection=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount :"
         Height          =   210
         Index           =   4
         Left            =   4935
         TabIndex        =   24
         Top             =   6180
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Net Amount :"
         Height          =   210
         Index           =   3
         Left            =   8850
         TabIndex        =   22
         Top             =   6180
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Balance Amount :"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   6195
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bank Code :"
         Height          =   210
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   555
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   375
         TabIndex        =   12
         Top             =   930
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Trans Code :"
         Height          =   255
         Left            =   75
         TabIndex        =   7
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "GRN Date :"
         Height          =   255
         Index           =   0
         Left            =   2985
         TabIndex        =   3
         ToolTipText     =   "Enter Value Date"
         Top             =   165
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            Description     =   "Add"
            Object.ToolTipText     =   "Add new record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            Description     =   "Edit"
            Object.ToolTipText     =   "Edit an existing record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            Description     =   "Remove "
            Object.ToolTipText     =   "Remove an existing record."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            Description     =   "Save a new Record"
            Object.ToolTipText     =   "Save on disk"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Slip"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
            Description     =   "Find a Record."
            Object.ToolTipText     =   "Find a record."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel"
            Description     =   "Cancel Operation"
            Object.ToolTipText     =   "Cancel operation mode"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   14
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4920
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPaymentsBank.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPaymentsBank.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPaymentsBank.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPaymentsBank.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPaymentsBank.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPaymentsBank.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPaymentsBank.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu File_Menu 
      Caption         =   "File"
      Begin VB.Menu New_Record 
         Caption         =   "New Record"
         Shortcut        =   ^N
      End
      Begin VB.Menu Edit_Record 
         Caption         =   "Edit Record"
         Shortcut        =   ^E
      End
      Begin VB.Menu Delete_Record 
         Caption         =   "Delete Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu Save_Record 
         Caption         =   "Save Record"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Edit_menu 
      Caption         =   "Edit"
      Begin VB.Menu Copy_date 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste_data 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Insert_row 
         Caption         =   "Insert Row"
         Shortcut        =   ^I
      End
      Begin VB.Menu Delete_row 
         Caption         =   "Delete Row"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmPOVendorPaymentsBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGRN As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_GRNTrans As New Recordset
Dim PR_ICIssue As New Recordset
Dim pr_dumy As New Recordset
Dim pr_dumy1 As New Recordset
Dim ls_sql As String
Dim ls_vno As String
Dim ls_BranchName As String
Dim ls_VchrDesc As String
Dim res, res1

Dim ClickRow
Private Sub TotalAmount()
    Dim ln_cnt As Integer
      txttotalamount = ""
       txtNetAmount = ""
      
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txtNetAmount = Val(txtNetAmount) + Val(.TextMatrix(ln_cnt, 16))
            txttotalamount = Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 4))
        Next
    End With
End Sub

Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<INV Code|<Sup Code|<Supplier Name|<Amount|<Excess Amt|<W/H Tax|<Instrument#|<TAmount|<GSTAmount|<SedAmount|<DiscAmount|<NetAmount|<Remarks|<GRNCCode|<Balance|<NetAmount|<Whtaxrate"
        .ColWidth(1) = 1200
        .ColWidth(2) = 1000
        .ColWidth(3) = 2700
        .ColWidth(4) = 1400
        .ColAlignment(4) = 7
        .ColWidth(5) = 1000
        .ColAlignment(5) = 7
        .ColWidth(6) = 1000
        .ColAlignment(6) = 7
        .ColWidth(7) = 1300
        
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 2700
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 1300
        .ColWidth(17) = 0
        .Redraw = True
    End With
    TxtRemarks = "CHQ. ISSUED TO SUPPLIERS"
    
End Sub
Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer)
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler

If KeyAscii = 13 Then 'if Enter Key then...
  Opt = ""
  With argFlexGrid
      '  .Row = .RowSel
       If .Col = 1 Then
       .CellBackColor = vbWindowBackground
       If .TextMatrix(.Row, 1) <> "" Then
       .TextMatrix(.Row, 1) = DoPad(.TextMatrix(.Row, 1), 10)
       .TextMatrix(.Row, 0) = .Row
       If SearchInGrid(GrdGRN, .TextMatrix(.Row, 1)) Then
         Call MsgBox("INV Already exist in Grid !!!", vbCritical)
        .CellBackColor = vbHighlight
        .Col = 1
        Exit Sub
       End If
       
       '--
       ls_sql = " SELECT  PO_POGRN.transcode, PO_POGRN.GRNCode,PO_POGRN.TransCode, PO_POGRN.AccountCode,PO_POGRN.flatdisc,PO_POGRN.netamount, IC_Supplier.Description,PO_POGRN.whstatus,PO_POGRN.whtaxrate FROM PO_POGRN INNER JOIN"
       ls_sql = ls_sql & " IC_Supplier ON PO_POGRN.Compcode = IC_Supplier.Compcode AND PO_POGRN.AccountCode = IC_Supplier.SupplierCode"
       ls_sql = ls_sql & " WHERE PO_POGRN.Compcode = '" & Gs_compcode & "' AND PO_POGRN.NetAmount <> PO_POGRN.RecAmount and PO_POGRN.transcode = '" & .TextMatrix(.Row, 1) & "'  "
       If pr_dumy.State = 1 Then pr_dumy.Close
       pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
       If Not pr_dumy.EOF Then
                .TextMatrix(.Row, 14) = Trim(pr_dumy("Grncode"))
                .TextMatrix(.Row, 2) = Trim(pr_dumy("Accountcode"))
                .TextMatrix(.Row, 3) = Trim(pr_dumy("Description") & "")
'                If pr_dumy1.State = 1 Then pr_dumy1.Close
'                 pr_dumy1.Open "SELECT SUM(Amount) as Amount , SUM(GSTAmount) as GSTAmount,SUM(SEDAmount) as SEDAmount , SUM(DiscAmount)  as DiscAmount   from  PO_POGRNDetail where compcode = '" & Gs_compcode & "' and transCode = '" & .TextMatrix(.Row, 1) & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'
'                 If Not pr_dumy1.EOF Then
'                 .TextMatrix(.Row, 8) = Val(pr_dumy1("Amount"))
'                 .TextMatrix(.Row, 9) = Val(pr_dumy1("GSTAmount"))
'                 .TextMatrix(.Row, 10) = Val(pr_dumy1("SEDAmount"))
'                 .TextMatrix(.Row, 11) = Val(pr_dumy1("DiscAmount"))
'                 End If
'                 pr_dumy1.Close
                 If Val(pr_dumy("whstatus")) = 1 Then
                  .TextMatrix(.Row, 6) = Round(((Val(pr_dumy("NetAmount")) * Val(pr_dumy("whtaxrate"))) / 100), 0)
                  .TextMatrix(.Row, 4) = Val(pr_dumy("NetAmount")) - Val(.TextMatrix(.Row, 6))
                  .TextMatrix(.Row, 16) = Val(pr_dumy("NetAmount"))
                  .TextMatrix(.Row, 17) = Val(pr_dumy("whtaxrate"))
                 Else
                  .TextMatrix(.Row, 4) = Val(pr_dumy("NetAmount"))
                  .TextMatrix(.Row, 16) = Val(pr_dumy("NetAmount"))
                  .TextMatrix(.Row, 17) = 0
                  
                 End If
                .TextMatrix(.Row, 12) = .TextMatrix(.Row, 4)
                .TextMatrix(.Row, 13) = "CHQ ISSUED PV #" + str(Val(.TextMatrix(.Row, 1)))
                .TextMatrix(.Row, 15) = Val(.TextMatrix(.Row, 4)) - CheckBalAmount(.TextMatrix(.Row, 1))
                 txtbalanceAmt = Val(.TextMatrix(.Row, 14))
                 If .Row = 1 Then
                 .TextMatrix(.Row, 7) = .TextMatrix(.Row, 7)
                 ElseIf .Rows > 2 Then
                 .TextMatrix(.Row, 7) = .TextMatrix(.Row - 1, 7)
                 End If
                  TotalAmount
                .Col = 4
            .CellBackColor = vbHighlight
      Else
            Call MsgBox("Invoice No not found !!!", vbCritical)
            .CellBackColor = vbHighlight
            .Col = 1
      End If
       pr_dumy.Close
      Else
        Call GrdGRN_KeyDown(112, vbKeyShift)
      End If

       ElseIf .Col = 2 Then
       ElseIf .Col = 3 Then
       ElseIf .Col = 4 Then
        .CellBackColor = vbWindowBackground
         If .TextMatrix(Row, 4) = "" Then
             Call MsgBox("Enter Amount!!!", vbCritical)
             Exit Sub
           End If
        .Col = .Col + 1
        .CellBackColor = vbHighlight
       ElseIf .Col = 5 Then
       .CellBackColor = vbWindowBackground
       
       .Col = .Col + 1
       .CellBackColor = vbHighlight
       ElseIf .Col = 6 Then
       .CellBackColor = vbWindowBackground
       
       .Col = .Col + 1
       .CellBackColor = vbHighlight
       
       ElseIf .Col = 7 Then
       .CellBackColor = vbWindowBackground
        .Col = 13
        TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
        TXTBARCODE.Text = .TextMatrix(.Row, 13)
        TXTBARCODE.Visible = True
        ClickRow = .Row
        TXTBARCODE.SetFocus
       ElseIf .Col = 7 Then
      
      If .TextMatrix(.Row, 1) <> "" Then
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .Row = .Row + 1
          .SetFocus
        Else
         Call MsgBox("Enter/Select GRN Code!!!", vbCritical)
         .Row = .Row
         .Col = 1
        End If
          
      
        If .RowSel > 9 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
        
        
            
   End If
   End With
 Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
    If .Col = 1 Or .Col = 4 Or .Col = 5 Or .Col = 6 Or .Col = 7 Then
      .CellBackColor = vbWindowBackground
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
      Opt = ""
      
      If .Col = 6 And Val(.TextMatrix(.Row, 17)) > 0 Then
          .TextMatrix(.Row, 4) = Val(.TextMatrix(.Row, 4)) - Val(.TextMatrix(.Row, 6))
      End If
      If .Col = 4 Or .Col = 5 Or .Col = 6 Then
          .TextMatrix(.Row, 16) = (Val(.TextMatrix(.Row, 4)) + Val(.TextMatrix(.Row, 5))) - Val(.TextMatrix(.Row, 6))
      End If
      TotalAmount
      
    End If
   End If
End With
End If

  If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      If .Col = 4 Or .Col = 5 Or .Col = 6 Then
         If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
         End If
         .Text = .Text & Chr(KeyAscii)
          If Not IsNumeric(.Text) Then
          .Text = ""
           Call MsgBox("Enter Numeric entry !!!", vbCritical)
          End If
          If .Col = 6 And Val(.TextMatrix(.Row, 17)) > 0 Then
          .TextMatrix(.Row, 4) = Val(.TextMatrix(.Row, 4)) - Val(.TextMatrix(.Row, 6))
          End If
          .TextMatrix(.Row, 16) = (Val(.TextMatrix(.Row, 4)) + Val(.TextMatrix(.Row, 5))) - Val(.TextMatrix(.Row, 6))
          TotalAmount
      If .Col = 4 Then
       'If Val(.TextMatrix(.Row, 4)) > Val(.TextMatrix(.Row, 12)) Then
       'Call MsgBox("Paid amount not greater then balance Amount !!!", vbCritical)
       ' .TextMatrix(.Row, 4) = 0
       ' .CellBackColor = vbHighlight
       'End If
       TotalAmount
      End If
      ElseIf .Col = 1 Then
        
         If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
         End If
         .Text = .Text & Chr(KeyAscii)
      ElseIf .Col = 7 Then
             .CellBackColor = vbWindowBackground
             .Text = .Text & Chr(KeyAscii)
      ElseIf .Col = 13 Then
        TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
       .TextMatrix(.Row, 13) = Chr(KeyAscii)
        TXTBARCODE.Text = .TextMatrix(.Row, 13)
        TXTBARCODE.Visible = True
        TXTBARCODE.SelStart = Len(TXTBARCODE)
        ClickRow = .Row
        TXTBARCODE.SetFocus
         
      End If
    End With
  End If
End Sub
Private Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from PO_PayableMaster ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(Val(0 & pr_dumy("transcode"))) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function

Private Sub Command2_Click()
If Mode = "A" Then
    LoadGRNTrans
    TxtRemarks = "Cash Payment to Vendors"
End If
End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbankcode
    Set PO_DESC = txtbankdesc
    Gs_SQL = "Select  Bankcode 'Code' ,Bankname from  SysBanks"
    Gs_FindFld = "Bankname"
    Gs_OrderBy = "Order by bankname"
    
    MyLookupOLDB.Caption = "Clients"
    MyLookupOLDB.Show 1
    
    If Len(txtbankcode) > 0 Then txtbankcode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Copy_date_Click()
With GrdGRN
Clipboard.Clear
Clipboard.SetText .TextMatrix(.Row, .Col)
End With
End Sub

Private Sub Delete_record_Click()
        Mode = DentMode(Mode, 3, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
        InitializeGrid
        txttransno.Enabled = True
        Command1.Enabled = True
        txttransno.SetFocus

End Sub

Private Sub Delete_row_Click()
   With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
            ResetRowSRNO
            TotalAmount
    End With

End Sub
Private Sub ResetRowSRNO()
With GrdGRN
   For ln_cnt = 1 To .Rows - 1
    .TextMatrix(ln_cnt, 0) = ln_cnt
   Next
End With
End Sub

Private Sub Edit_record_Click()
        Mode = DentMode(Mode, 2, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
        InitializeGrid
        txttransno.Enabled = True
        Command1.Enabled = True
        txttransno.SetFocus
End Sub

Private Sub GrdGRN_Click()
GrdGRN.SelectionMode = flexSelectionFree
With GrdGRN
GrdGRN.ToolTipText = .TextMatrix(.Row, .Col)
End With
GrdGRN.CellBackColor = vbHighlight
If GrdGRN.Col = 12 Then
With GrdGRN

        TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
        TXTBARCODE.Text = .TextMatrix(.Row, 12)
        TXTBARCODE.Visible = True
        ClickRow = .Row
        TXTBARCODE.SetFocus
End With
End If
 txtbalanceAmt = Val(GrdGRN.TextMatrix(GrdGRN.Row, 14))
End Sub

Private Sub GrdGRN_DblClick()
    GrdGRN.SelectionMode = flexSelectionFree
End Sub

Private Sub GrdGRN_EnterCell()
GrdGRN.CellBackColor = vbHighlight
If GrdGRN.Col = 12 Then
With GrdGRN

        TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
        TXTBARCODE.Text = .TextMatrix(.Row, 12)
        TXTBARCODE.Visible = True
        ClickRow = .Row
        TXTBARCODE.SetFocus
End With
End If

End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then  ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    Gs_SQL = " SELECT PO_POGRN.transcode, IC_Supplier.Description,PO_POGRN.GRNCode,PO_POGRN.netamount"
    Gs_SQL = Gs_SQL & " FROM PO_POGRN INNER JOIN IC_Supplier ON PO_POGRN.Compcode = IC_Supplier.Compcode AND PO_POGRN.AccountCode = IC_Supplier.SupplierCode"
    Gs_OtherPara = " WHERE PO_POGRN.Compcode = '" & Gs_compcode & "' AND PO_POGRN.NetAmount <> PO_POGRN.RecAmount"
    Gs_FindFld = "GRNCode"
    MyLookupOLDBsearchmultipul.Caption = "Unpaid GRN"
    MyLookupOLDBsearchmultipul.Show 1
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = Text1
    If GrdGRN.TextMatrix(GrdGRN.Row, 1) <> "" Then
        Call GrdGRN_KeyPress(13)
    End If
 ElseIf KeyCode = vbKeyDelete Then  'Delete Key Pressed
            With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
             ResetRowSRNO
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
            TotalAmount
           End With
 
 ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then  'Key Pressed
 txtbalanceAmt = Val(GrdGRN.TextMatrix(GrdGRN.Row, 14))
 End If

 
End Sub

Private Sub GrdGRN_KeyPress(KeyAscii As Integer)
'On Error GoTo ErrHandler
 Call GetKeysAdd(GrdGRN, KeyAscii)
Exit Sub
'ErrHandler:
'MsgBox ("An Error has Occured In The MSFlexgrid1_KeyPress() Procedure") & vbCr & "Report This Error To Latifjat@hotmail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.Number & vbCr & "Error Description : " & Err.Description, vbCritical, "FlexGrid Example"
End Sub



Private Sub GrdGRN_LeaveCell()
With GrdGRN
    .CellBackColor = vbWindowBackground
End With
End Sub


Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttransno
    Set PO_DESC = Text1
    Gs_SQL = "Select TransCode, Transdate from PO_PayableMaster"
    Gs_FindFld = "TransCodebank"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' and GLstatus = 0 "
    Gs_OrderBy = "Order by TransCode"

    MyLookupOLDB.Caption = "Vendor Payments"
    MyLookupOLDB.Show 1
    
    If txttransno <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Form_Load()
  SetToolBar(1) = chkRights("VENDORPMT1")
  SetToolBar(2) = chkRights("VENDORPMT2")
  SetToolBar(3) = chkRights("VENDORPMT3")
  SetToolBar(4) = chkRights("VENDORPMT4")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  Me.Caption = Me.Caption + " (" + Gs_CompName + ")"
  txtvaluedate.Value = Date
  
 InitializeGrid

End Sub

Private Sub Insert_row_Click()
With GrdGRN
If .TextMatrix(.Row, 1) <> "" Then
          .CellBackColor = vbWindowBackground
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .LeftCol = 1
          .Row = .Row + 1
          .Row = .Rows - 1
          .SetFocus
        Else
         Call MsgBox("Enter/Select GRN Code!!!", vbCritical)
         .Row = .Row
         .Col = 1
        End If
          
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
End With
End Sub

Private Sub New_Record_Click()
        Mode = DentMode(Mode, 1, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
        txttransno = maxtranscode
        InitializeGrid
        txttransno.Enabled = False
        Command1.Enabled = False
        txtvaluedate.SetFocus
End Sub

Private Sub Paste_data_Click()
With GrdGRN
.TextMatrix(.Row, .Col) = Clipboard.GetText
End With
End Sub

Private Sub Save_Record_Click()
        Mode = DentMode(Mode, 4, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
       Mode = DentMode(Mode, Button.Index, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
       If Mode = "A" Then
        txttransno = maxtranscode
        InitializeGrid
        txttransno.Enabled = False
        Command1.Enabled = False
        txtvaluedate.SetFocus
        
       Else
        txttransno.Enabled = True
        Command1.Enabled = True
        txttransno.SetFocus
       End If
If Button.Index = 7 Then InitializeGrid
End Sub

Public Sub SaveValues()
'On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String
Dim ls_transCodecash As String



gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
                        gc_dbcon.Execute "DELETE FROM PO_PayableMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                        gc_dbcon.Execute "DELETE FROM PO_PayableDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              
                         With GrdGRN
                                For ln_cnt = 1 To .Rows - 1
                                    ls_sql = "Update  PO_POGRN set RecAmount = " & CheckBalAmount(Trim(.TextMatrix(ln_cnt, 1))) & "  WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(.TextMatrix(ln_cnt, 1)) & "'"
                                    gc_dbcon.Execute ls_sql
                                Next
                         End With
              
           Case Else
                If Mode = "E" Then
                          gc_dbcon.Execute "DELETE FROM PO_PayableMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                          gc_dbcon.Execute "DELETE FROM PO_PayableDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                End If
                If Mode = "A" Then
                    txttransno = maxtranscode
                End If
                        ls_sql = "INSERT into PO_PayableMaster( Compcode,branchcode, TransCode, TransDate,PaymentType,Remarks,userid,adddate,addtime,Bankcode)"
                        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "',1,'" & RepApp(TxtRemarks) & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "' ,'" & txtbankcode & "')"
                        gc_dbcon.Execute ls_sql
                
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      If .TextMatrix(ln_cnt, 1) <> "" And Val(.TextMatrix(ln_cnt, 4)) <> 0 Then
                
                        ls_sql = "INSERT into PO_PayableDetail(Compcode,BranchCode, TransCode, GRNCode, MGRNCode, Amount,GSTAmount,SedAmount,DiscAmount,PaidAmount,ExcessAmount,Remarks,whtax,instrumentno,accountcode,whtaxrate)"
                        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 14)) & "'," & Val(.TextMatrix(ln_cnt, 8)) & "," & Val(.TextMatrix(ln_cnt, 9)) & "," & Val(.TextMatrix(ln_cnt, 10)) & "," & Val(.TextMatrix(ln_cnt, 11)) & "," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & ",'" & Trim(.TextMatrix(ln_cnt, 13)) & "'," & Val(.TextMatrix(ln_cnt, 6)) & ",'" & Trim(.TextMatrix(ln_cnt, 7)) & "','" & Trim(.TextMatrix(ln_cnt, 2)) & "'," & Val(.TextMatrix(ln_cnt, 17)) & ")"
                        gc_dbcon.Execute ls_sql
                        
                        
                        ls_sql = "Update  PO_POGRN set RecAmount =  " & CheckBalAmount(Trim(.TextMatrix(ln_cnt, 1))) & " WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(.TextMatrix(ln_cnt, 1)) & "'"
                        gc_dbcon.Execute ls_sql
                     End If
                    Next
                 End With
                 
                 
                 
     End Select
gc_dbcon.CommitTrans

If Mode = "A" Or Mode = "E" Then
    res = MsgBox("Do you want to post GL Voucher", vbYesNo + vbInformation)
    If res = vbYes Then
         Call PostPayableVoucherBank(Gs_compcode, txtvaluedate, txtvaluedate, txttransno)
         
         res1 = MsgBox("Do you want to Print Voucher", vbYesNo + vbInformation)
         If res1 = vbYes Then
            If pr_dumy1.State = 1 Then pr_dumy1.Close
            pr_dumy1.Open "select voucherno from PO_PayableMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            
            If Not pr_dumy1.EOF Then
            ls_vno = Trim(pr_dumy1("VoucherNo") & "")
            End If
            pr_dumy1.Close
            Call setprint
            If txtbankcode.Enabled Then txtbankcode.SetFocus
            
         End If
    End If

End If


If Mode = "A" Then
    txttransno = maxtranscode
End If

InitializeGrid
Exit Sub
RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub setprint()
On Error GoTo LocalErr

If ls_vno <> "" Then

If pr_dumy1.State = 1 Then pr_dumy1.Close
pr_dumy1.Open "Select VchrDescrip from GlVchrType where vchrtype = 'BPP' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
    ls_VchrDesc = Trim(pr_dumy1("VchrDescrip") & "")
    ls_BranchName = "Head Office"
End If
pr_dumy1.Close
   
   
   
   
   With rptVoucher
        
        .ReportFileName = App.Path & Gs_GlRepoPath & "\Vchr_Print.RPT"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & ls_VchrDesc & "'"
        .Formulas(5) = "BranchName = '" & Gs_BranchCode + "-" + ls_BranchName & "'"
        .SelectionFormula = "{Gl_Trans.Voucher_No} = '" & Trim(ls_vno) & "' and {Gl_Trans.BranchCode} = '" & Gs_BranchCode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.VchrType} = 'BPP'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.CompCode} = '" & Gs_compcode & "'"
        '.SelectionFormula = .SelectionFormula & " and {Gl_Trans.Value_Date} = Date(" & Year(txtvaluedate) & "," & Month(txtvaluedate) & "," & Day(txtvaluedate) & ")"
        .Formulas(2) = "Sig1 = '" & Gc_UserName & "'"
        .Formulas(3) = "Sig2 = '" & Gs_Sign2 & "'"
        .Formulas(4) = "Sig3 = '" & Gs_Sign3 & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
End If

Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub

Private Function CheckBalAmount(ls_GrnCODE As String) As Double
Dim ln_PaidAmount As Double


pr_dumy1.Open "Select Sum(paidAmount)+Sum(whtax) as Amount from  PO_PayableDetail where compcode = '" & Gs_compcode & "' and GRNCode = '" & ls_GrnCODE & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1

If Not pr_dumy1.EOF Then
    ln_PaidAmount = Val(0 & pr_dumy1("Amount"))
End If
pr_dumy1.Close

CheckBalAmount = ln_PaidAmount

End Function
Public Sub ClearVal()
     '
End Sub

Private Sub SetVal()
     txtvaluedate = PR_ICIssue("Transdate")
     TxtRemarks = Trim(PR_ICIssue("Remarks") & "")
     txtvchrType = Trim(PR_ICIssue("VchrType") & "")
     txtvchrno = Trim(PR_ICIssue("VoucherNo") & "")
     txtbankcode = Trim(PR_ICIssue("Bankcode") & "")
     If txtbankcode <> "" Then Call txtbankcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Public Function ChkInputs() As Boolean
    If Trim(txttransno.Text) = "" Then
      Call MsgBox("Enter/Select Payment ID !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(TxtRemarks.Text) = "" Then
      Call MsgBox("Enter Remarks !!!", vbCritical)
      ChkInputs = False
    ElseIf GrdGRN.TextMatrix(1, 1) = "" Then
      Call MsgBox("Enter Grid Entries !!!", vbCritical)
      ChkInputs = False
    Else
       ChkInputs = True
    End If
End Function
Private Sub LoadPaidTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String
ls_sql = " SELECT PO_PayableDetail.MGRNCode,PO_PayableDetail.instrumentno,PO_PayableDetail.GRNCode,PO_PayableDetail.whtax,PO_Payabledetail.AccountCode,IC_Supplier.Description, PO_PayableDetail.excessAmount , PO_PayableDetail.paidAmount, PO_PayableDetail.Amount, PO_PayableDetail.GSTAmount, PO_PayableDetail.SEDAmount, "
ls_sql = ls_sql & " PO_PayableDetail.DiscAmount , PO_PayableDetail.Remarks FROM  PO_PayableMaster INNER JOIN"
ls_sql = ls_sql & " PO_PayableDetail ON PO_PayableMaster.Compcode = PO_PayableDetail.Compcode AND"
ls_sql = ls_sql & " PO_PayableMaster.BranchCode = PO_PayableDetail.BranchCode And PO_PayableMaster.TransCode = PO_PayableDetail.TransCode INNER JOIN"
ls_sql = ls_sql & " IC_Supplier ON PO_PayableMaster.Compcode = IC_Supplier.Compcode AND PO_Payabledetail.AccountCode = IC_Supplier.SupplierCode "
ls_sql = ls_sql & "  where PO_PayableMaster.Compcode = '" & Gs_compcode & "' and PO_PayableMaster.Transcode = '" & txttransno & "' and PO_PayableMaster.glstatus = 0 order by PO_PayableDetail.srno"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                 
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("GRNCode") & "")
                .TextMatrix(.Row, 14) = Trim(Pr_LoadTrans("MGRNCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("AccountCode") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 8) = Val(Pr_LoadTrans("Amount"))
                .TextMatrix(.Row, 5) = Val(Pr_LoadTrans("excessAmount"))
                .TextMatrix(.Row, 9) = Val(Pr_LoadTrans("GSTAmount"))
                .TextMatrix(.Row, 10) = Val(Pr_LoadTrans("SEDAmount"))
                .TextMatrix(.Row, 11) = Val(Pr_LoadTrans("DiscAmount"))
                .TextMatrix(.Row, 4) = Val(Pr_LoadTrans("PaidAmount"))
                .TextMatrix(.Row, 6) = Val(Pr_LoadTrans("whtax"))
                .TextMatrix(.Row, 15) = (Val(.TextMatrix(.Row, 8)) + Val(.TextMatrix(.Row, 9)) + Val(.TextMatrix(.Row, 10))) - (Val(.TextMatrix(.Row, 11))) - CheckBalAmount(.TextMatrix(.Row, 14))
                .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 4)) + Val(.TextMatrix(.Row, 15))
                .TextMatrix(.Row, 7) = Trim(Pr_LoadTrans("instrumentno") & "")
                .TextMatrix(.Row, 13) = Trim(Pr_LoadTrans("Remarks") & "")
                .TextMatrix(.Row, 16) = Val(.TextMatrix(.Row, 4)) + Val(.TextMatrix(.Row, 5)) + Val(.TextMatrix(.Row, 6))
                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalAmount
    Else
        Call SetErr("Transaction not found.!!!", vbCritical)
        
    End If
    Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub


Private Sub txtbankcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtbankcode <> "" Then
    txtbankcode = DoPad(txtbankcode, txtbankcode.MaxLength)
    
        ls_sql = "Select bankcode,bankname as Description from SysBanks where bankcode = '" & txtbankcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Bank code not found", vbCritical)
                'Cancel = True
            Else
                txtbankdesc = pr_dumy("description")
                If TxtRemarks.Enabled Then TxtRemarks.SetFocus
            End If
         pr_dumy.Close
ElseIf KeyCode = vbKeyReturn And txtbankcode = "" Then
        Command3_Click
End If
End Sub

Private Sub TXTBARCODE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyTab Then
TXTBARCODE_LostFocus
End If
End Sub

Private Sub TXTBARCODE_LostFocus()
With GrdGRN
If ClickRow <> "" Then
.TextMatrix(ClickRow, 13) = TXTBARCODE.Text
 TXTBARCODE.Text = ""
 ClickRow = ""
 End If
 .CellBackColor = vbWindowBackground
 TXTBARCODE.Visible = False
 If .TextMatrix(.Row, 1) <> "" Then
   If .Row = .Rows - 1 Then
     .Rows = .Rows + 1
   End If
 .Row = .Row + 1
 .SetFocus
 .Col = 1
 .Row = .Row
 .LeftCol = 1
 
 End If
  If .RowSel > 9 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
  End If
        
End With


End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then GrdGRN.SetFocus
End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)


 If KeyCode = vbKeyReturn And Len(txttransno.Text) > 0 Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txttransno.Text = DoPad(UCase(txttransno.Text), 10)
         PR_ICIssue.Open "select * from PO_PayableMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       Select Case Mode
            Case "A"
                If Not PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   If txttransno.Enabled Then txttransno.SetFocus
                Else
                   txtvaluedate.SetFocus
                End If
            Case Else
                If PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txttransno.SetFocus
                Else
                   Call SetVal
                   LoadPaidTrans
                   If Mode <> "D" Then
                      txttransno.SetFocus
                   End If
                End If
            End Select
     ElseIf KeyCode = vbKeyF12 Then
           ' cmdLookup_Click
     End If
  End Sub
Private Sub LoadGRNTrans()
Dim PR_GRNTrans As New Recordset
Dim pr_dumy1 As New Recordset
Dim ls_Exp As String
Dim ls_sql As String
InitializeGrid



ls_sql = " SELECT PO_POGRN.GRNcode,PO_POGRN.TransCode, PO_POGRN.AccountCode,PO_POGRN.flatdisc, IC_Supplier.Description FROM PO_POGRN INNER JOIN"
ls_sql = ls_sql & " IC_Supplier ON PO_POGRN.Compcode = IC_Supplier.Compcode AND PO_POGRN.AccountCode = IC_Supplier.SupplierCode"
ls_sql = ls_sql & " WHERE PO_POGRN.Compcode = '" & Gs_compcode & "' AND PO_POGRN.NetAmount <> PO_POGRN.RecAmount and PO_POGRN.Transdate = '" & Format(txtvaluedate, "YYYY/MM/DD") & "'  "

PR_GRNTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_GRNTrans.EOF Then
        With GrdGRN
            Do While Not PR_GRNTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_GRNTrans("TransCode") & "")
                .TextMatrix(.Row, 14) = Trim(PR_GRNTrans("GrnCode") & "")
                .TextMatrix(.Row, 2) = Trim(PR_GRNTrans("Accountcode"))
                .TextMatrix(.Row, 3) = Trim(PR_GRNTrans("Description") & "")
                 pr_dumy1.Open "SELECT SUM(Amount) as Amount , SUM(GSTAmount) as GSTAmount,SUM(SEDAmount) as SEDAmount , SUM(DiscAmount) as DiscAmount   from  PO_POGRNDetail where compcode = '" & Gs_compcode & "' and transCode = '" & .TextMatrix(.Row, 14) & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1

                 If Not pr_dumy1.EOF Then
                 .TextMatrix(.Row, 8) = Val(pr_dumy1("Amount"))
                 .TextMatrix(.Row, 9) = Val(pr_dumy1("GSTAmount"))
                 .TextMatrix(.Row, 10) = Val(pr_dumy1("SEDAmount"))
                 .TextMatrix(.Row, 11) = Val(pr_dumy1("DiscAmount"))
                 
                 
                 End If
                 pr_dumy1.Close

                
                .TextMatrix(.Row, 4) = (Val(.TextMatrix(.Row, 8)) + Val(.TextMatrix(.Row, 9)) + Val(.TextMatrix(.Row, 10))) - (Val(.TextMatrix(.Row, 11)) + Val(0 & PR_GRNTrans("flatdisc")))
                .TextMatrix(.Row, 4) = Val(.TextMatrix(.Row, 4)) - CheckBalAmount(.TextMatrix(.Row, 14))
                .TextMatrix(.Row, 12) = .TextMatrix(.Row, 4)
                .TextMatrix(.Row, 13) = "Amount paid to " + .TextMatrix(.Row, 3) + " Payment # =" + txttransno
                .TextMatrix(.Row, 14) = Val(.TextMatrix(.Row, 4))
                .TextMatrix(.Row, 16) = Val(.TextMatrix(.Row, 4)) + Val(.TextMatrix(.Row, 5)) + Val(.TextMatrix(.Row, 6))
                .TextMatrix(.Row, 17) = Val(pr_dumy1("whtaxrate"))
                .Rows = .Rows + 1

                PR_GRNTrans.MoveNext
                If PR_GRNTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalAmount
        
    Else
        Call SetErr("Transaction not found.!!!", vbCritical)
        
    End If
    PR_GRNTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub
Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then txtbankcode.SetFocus
End Sub

