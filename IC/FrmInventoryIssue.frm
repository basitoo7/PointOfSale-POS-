VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInventoryIssue 
   Caption         =   "Stock Issue Note"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInventoryIssue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   12345
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
      Height          =   1590
      Left            =   60
      TabIndex        =   1
      Top             =   570
      Width           =   12240
      Begin VB.TextBox txtGrndesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4995
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   150
         Width           =   4035
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   4665
         Picture         =   "FrmInventoryIssue.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   135
         Width           =   315
      End
      Begin VB.TextBox txtGrnNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   3570
         MaxLength       =   10
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1500
         Picture         =   "FrmInventoryIssue.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1140
         Width           =   315
      End
      Begin VB.TextBox txtacode 
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
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   1155
         Width           =   435
      End
      Begin VB.TextBox txtaname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1140
         Width           =   4035
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   7470
         Picture         =   "FrmInventoryIssue.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1155
         Width           =   315
      End
      Begin VB.TextBox txtacode1 
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
         Left            =   7020
         MaxLength       =   3
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   1170
         Width           =   435
      End
      Begin VB.TextBox txtaname1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1155
         Width           =   4290
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
         Height          =   315
         Left            =   1050
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2160
         Picture         =   "FrmInventoryIssue.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   150
         Width           =   315
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1110
         Visible         =   0   'False
         Width           =   195
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   10650
         TabIndex        =   9
         Top             =   165
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63176705
         CurrentDate     =   37580
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3870
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1140
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   510
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   570
         Width           =   11025
      End
      Begin Crystal.CrystalReport rptVoucher 
         Left            =   10560
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
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
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7470
         Top             =   -180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "GRN #  :"
         Height          =   255
         Left            =   2295
         TabIndex        =   28
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Verified By :"
         Height          =   255
         Left            =   75
         TabIndex        =   24
         Top             =   1185
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Approved By :"
         Height          =   255
         Left            =   5955
         TabIndex        =   23
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "SIN #  :"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   195
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   555
         Width           =   780
      End
      Begin VB.Label label2 
         Caption         =   "Issue Date :"
         Height          =   255
         Left            =   9780
         TabIndex        =   5
         ToolTipText     =   "Enter Value Date"
         Top             =   180
         Width           =   930
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   1058
      ButtonWidth     =   1402
      ButtonHeight    =   1005
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
               Picture         =   "FrmInventoryIssue.frx":08D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryIssue.frx":0D26
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryIssue.frx":117A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryIssue.frx":15CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryIssue.frx":1A22
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryIssue.frx":1E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryIssue.frx":25CA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5310
      Left            =   45
      TabIndex        =   2
      Top             =   2055
      Width           =   12255
      Begin VB.TextBox txttotalamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   10800
         MaxLength       =   11
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   4905
         Width           =   1320
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   4650
         Left            =   150
         TabIndex        =   16
         Top             =   195
         Width           =   11970
         _ExtentX        =   21114
         _ExtentY        =   8202
         _Version        =   393216
         RowHeightMin    =   300
         BackColorSel    =   16777215
         ForeColorSel    =   0
         GridColor       =   8421504
         AllowUserResizing=   1
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
      Begin VB.TextBox txtitemname 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   4935
         Width           =   9090
      End
      Begin VB.TextBox txtnoofitems 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Qty:"
         Height          =   255
         Left            =   9690
         TabIndex        =   25
         Top             =   4950
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   " Total :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6345
         TabIndex        =   7
         Top             =   3180
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu New_Record 
         Caption         =   "&New Record"
         Shortcut        =   ^N
      End
      Begin VB.Menu edit_Record 
         Caption         =   "&Edit Record"
         Shortcut        =   ^E
      End
      Begin VB.Menu delete_Record 
         Caption         =   "&Delete Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu save_Record 
         Caption         =   "&Save Record"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "Edit"
      Begin VB.Menu copy_data 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste_data 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu AddNewRow 
         Caption         =   "Add New Row"
         Shortcut        =   ^I
      End
      Begin VB.Menu DeleteRow 
         Caption         =   "Delete Row"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmInventoryIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGRN As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object

Dim ln_cnt As Integer
Dim Resopt

Dim Po_Status  As Integer
Dim Ls_ItemName  As String
Dim ln_qty, LN_EnterQty
Dim ls_transtype As String

Dim pr_dumy As New Recordset

Dim PR_UOM As New Recordset

Dim PR_ICIssue As New Recordset

Dim PR_IcItem As New Recordset
Dim PR_Branch As New Recordset
Dim ln_res
Private Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from IC_IssueNoteMaster where compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
txtCustRef = ClientCoderef("015") + Right(maxtranscode, 4)
End Function



Private Sub AddNewRow_Click()
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
         Call MsgBox("Enter/Select Item Code!!!", vbCritical)
         .Row = .Row
         .Col = 1
        End If
          
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
End With
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtacode
    Set PO_DESC = txtaname
    Gs_SQL = "Select ACode, Aname Description from PO_AuthorityPerson "
    Gs_FindFld = "Aname"
    Gs_OrderBy = "Order by AName"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Authority Person"
    MyLookupOLDB.Show 1
    
    If txtacode <> "" Then Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command4_Click()
        On Error Resume Next
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txtGRNNo
        Set PO_DESC = txtGrndesc
        Gs_SQL = "SELECT GRN.TransCode AS ComputerCode,  Vendors.Description AS 'Vendors.Description',GRN.GRNCode AS GRNCode, GRN.TransDate AS GRNDate,    GRN.NetAmount AS 'GRN.NetAmount' FROM         PO_POGRN GRN INNER JOIN         IC_Supplier Vendors ON GRN.Compcode = Vendors.Compcode AND GRN.AccountCode = Vendors.SupplierCode"
        Gs_OrderBy = "ORDER BY GRN.TransCode desc"
        Gs_OtherPara = " Where GRN.compcode = '" & Gs_compcode & "' and GRN.glstatus = 0 "
        
        frmPosearchRecords.Caption = "GRN"
        frmPosearchRecords.Show 1
        
        If txtGRNNo <> "" Then Call txtGRNNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub txtGRNNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Trim(txtGRNNo.Text) <> "" Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txtGRNNo.Text = DoPad(UCase(txtGRNNo.Text), 10)
         PR_ICIssue.Open "select * from PO_POGRN where compcode = '" & Gs_compcode & "' and Transcode = '" & txtGRNNo & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         
              If PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   If txtGRNNo.Enabled Then txtGRNNo.SetFocus
                    txtGRNNo.Enabled = True
                 
                Else
                   LoadGRNTrans1
                End If
            
 ElseIf KeyCode = vbKeyReturn And Trim(txtGRNNo.Text) = "" Then
           Command4_Click
 End If
 End Sub
Private Sub LoadGRNTrans1()
On Error GoTo LocalErr

Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String


       

ls_sql = " SELECT PO_POGRNDetail.CustomCode,PO_POGRNDetail.ItemCode,PO_POGRNDetail.siteid, IC_Item.salecost,IC_Item.Description, PO_POGRNDetail.Quantity, PO_POGRNDetail.Rate, PO_POGRNDetail.Amount, PO_POGRNDetail.BonusQty, PO_POGRNDetail.BonusAmount ,PO_POGRNDetail.GSTper,PO_POGRNDetail.Sedper,PO_POGRNDetail.GSTAmount,PO_POGRNDetail.SedAmount,PO_POGRNDetail.Discper,PO_POGRNDetail.Discamount,PO_POGRNDetail.Remarks, IC_ItemUM.Description AS UOM"
ls_sql = ls_sql & " ,PO_POGRNDetail.expdays,PO_POGRNDetail.expdate,PO_POGRNDetail.FlatAmount FROM PO_POGRNDetail INNER JOIN IC_Item ON PO_POGRNDetail.Compcode = IC_Item.Compcode AND PO_POGRNDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where PO_POGRNDetail.Compcode = '" & Gs_compcode & "' and PO_POGRNDetail.Transcode = '" & txtGRNNo & "' order by PO_POGRNDetail.srno"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                
                
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("CustomCode") & "")
                .TextMatrix(.Row, 8) = Trim(Pr_LoadTrans("ItemCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("UOM") & "")
                .TextMatrix(.Row, 4) = CheckBalQTY(.TextMatrix(.Row, 8), 1)
                .TextMatrix(.Row, 6) = Pr_LoadTrans("Quantity")
                .TextMatrix(.Row, 5) = Val(0 & Pr_LoadTrans("Rate"))
                .TextMatrix(.Row, 9) = Val(0 & Pr_LoadTrans("Rate"))
                .TextMatrix(.Row, 7) = Trim(Pr_LoadTrans("Remarks") & "")
           
                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalAmount
    Else
        Call SetErr("GRN Transaction not found.!!!", vbCritical)
        
    End If
    Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub


Private Sub Copy_data_Click()
With GrdGRN
Clipboard.Clear
Clipboard.SetText .TextMatrix(.Row, .Col)
End With
End Sub

Private Sub Delete_record_Click()
   txttransno.Enabled = True
   txttransno.SetFocus
   Command1.Enabled = True
   Mode = DentMode(Mode, 3, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
End Sub

Private Sub DeleteRow_Click()
   With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
            ResetRowSRNO
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
   txttransno.Enabled = True
   txttransno.SetFocus
   Command1.Enabled = True
   Mode = DentMode(Mode, 2, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
   CheckLogTrans
End Sub

Private Sub New_Record_Click()
       txttransno.Enabled = False
       Command1.Enabled = False
       InitializeGrid
       txttransno = maxtranscode
       Mode = DentMode(Mode, 1, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
       TxtRemarks = "Stock Issue to Showroom"
       CheckLogTrans
       TxtRemarks.SetFocus
End Sub

Private Sub Paste_data_Click()
With GrdGRN
.TextMatrix(.Row, .Col) = Clipboard.GetText
End With
End Sub

Private Sub Save_Record_Click()
   Mode = DentMode(Mode, 4, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
End Sub

Private Sub txtACode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtacode) <> "" And KeyCode = vbKeyReturn Then
        txtacode = DoPad(txtacode, 3)
        pr_dumy.Open "Select * from PO_AuthorityPerson where Compcode  = '" & Gs_compcode & "' and Acode = '" & txtacode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Authority Code not found !!!", vbCritical)
            txtacode = ""
            txtaname = ""
            txtacode.SetFocus
        Else
            txtaname = pr_dumy("aname")
            txtacode1.SetFocus
        End If
        pr_dumy.Close

ElseIf Trim(txtacode) = "" And KeyCode = vbKeyReturn Then
        txtacode = ""
        txtaname = ""
End If

End Sub


Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtacode1
    Set PO_DESC = txtaname1
    Gs_SQL = "Select ACode, Aname Description from PO_AuthorityPerson "
    Gs_FindFld = "Aname"
    Gs_OrderBy = "Order by AName"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Authority Person"
    MyLookupOLDB.Show 1
    
    If txtacode1 <> "" Then Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)

End Sub
Private Sub txtACode1_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtacode1) <> "" And KeyCode = vbKeyReturn Then
        txtacode1 = DoPad(txtacode1, 3)
        pr_dumy.Open "Select * from PO_AuthorityPerson where Compcode  = '" & Gs_compcode & "' and Acode = '" & txtacode1 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Authority Code not found !!!", vbCritical)
            txtacode1 = ""
            txtaname1 = ""
            txtacode1.SetFocus
        Else
            txtaname1 = pr_dumy("aname")
            GrdGRN.SetFocus
        End If
        pr_dumy.Close

ElseIf Trim(txtacode1) = "" And KeyCode = vbKeyReturn Then
        txtacode1 = ""
        txtaname1 = ""
End If
End Sub


Private Sub GrdGRN_EnterCell()
GrdGRN.CellBackColor = vbHighlight
End Sub

Private Sub GrdGRN_LeaveCell()
With GrdGRN
    .CellBackColor = vbWindowBackground
End With
GrdGRN.SelectionMode = flexSelectionFree
End Sub
Private Sub GrdGRN_Click()
GrdGRN.SelectionMode = flexSelectionFree
With GrdGRN
   txtitemname = .TextMatrix(.Row, 2)
End With
GrdGRN.CellBackColor = vbHighlight
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttransno
    Set PO_DESC = Text1
    Gs_SQL = "Select TransCode, Transdate from IC_IssueNoteMaster "
    Gs_FindFld = "TransCode"
    Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "'  and glstatus = 0"
    Gs_OrderBy = "Order by TransCode desc"

    MyLookupOLDB.Caption = "Stock Issue Note"
    MyLookupOLDB.Show 1
    
    If txttransno <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_ICIssue, Me, txttransno, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
  ls_transtype = "D"
  SetToolBar(1) = chkRights("INVISSUE01")
  SetToolBar(2) = chkRights1("INVIS00002")
  SetToolBar(3) = chkRights1("INVIS00003")
  SetToolBar(4) = chkRights("INVISSUE04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  

  txtvaluedate.Value = Date
 
  InitializeGrid
  

  
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtacode.SetFocus
    txtacode = "001"
    txtacode1 = "001"
    If txtacode <> "" Then Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
    If txtacode1 <> "" Then Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)
End If
End Sub
Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)


 If KeyCode = vbKeyReturn And Len(txttransno.Text) > 0 Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txttransno.Text = DoPad(UCase(txttransno.Text), 10)
         PR_ICIssue.Open "select * from IC_IssueNoteMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0 ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
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
                   txttransno.Enabled = True
                   txttransno.SetFocus
                
                Else
                   Call SetVal
                   LoadGRNTrans
                   txttransno.Enabled = False
                   
                End If
            End Select
     ElseIf KeyCode = vbKeyF12 Then
            
     End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       txttransno.Enabled = False
       Command1.Enabled = False
       InitializeGrid
       
    Else
      txttransno.Enabled = True
       txttransno.SetFocus
       Command1.Enabled = True
    End If
    If Button.Index = 7 Then
    InitializeGrid
    End If
    
    If PB_BlnkGRN And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
       txttransno = maxtranscode
       TxtRemarks = "Stock Issue to Showroom"
       TxtRemarks.SetFocus
       txttransno.Enabled = False
       Command1.Enabled = False
       CheckLogTrans
    End If
    
End Sub
Private Sub CheckLogTrans()
Dim pr_dumyLog As New Recordset
Dim res
pr_dumyLog.Open "select * from IC_IssueNoteDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyLog.EOF Then
    If pr_dumyLog("Emode") = "E" Then
        txttransno = pr_dumyLog("Transcode")
        res = MsgBox(txttransno & " # you have opened in edit mode not save Do you want to open now", vbYesNo + vbExclamation)
        If res = vbYes Then
        Mode = DentMode(Mode, 2, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
        If PR_ICIssue.State = 1 Then PR_ICIssue.Close
        PR_ICIssue.Open "select * from IC_IssueNoteMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0 ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If Not PR_ICIssue.EOF Then
        Call SetVal
        End If
        PR_ICIssue.Close
        LoadLogTrans
        Else
        
        ls_sql = "delete from  IC_IssueNoteDetaillog where computername = '" & Gs_ComputerName & "' "
        gc_dbcon.Execute ls_sql
           
        End If
    Else
        LoadLogTrans
    End If
End If
pr_dumyLog.Close
End Sub

Public Sub SaveValues()
'On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String

gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
              gc_dbcon.Execute "DELETE FROM IC_IssueNoteMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              gc_dbcon.Execute "DELETE FROM IC_IssueNoteDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              
              
           Case Else
                If Mode = "E" Then
                    gc_dbcon.Execute "DELETE FROM IC_IssueNoteMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                    gc_dbcon.Execute "DELETE FROM IC_IssueNoteDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                End If
                If Mode = "A" Then
                    txttransno = maxtranscode
                    Me.Refresh
                End If
                    
                      ls_sql = "INSERT into IC_IssueNoteMaster( Compcode,branchcode, TransCode,PCode,   TransDate, AccountCode, SiteID, BinID, Remarks,vcode,acode,userid,adddate,addtime)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(txtGRNNo) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','000001','001','001','" & RepApp(TxtRemarks) & "','" & txtacode & "','" & txtacode1 & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "' )"
                      gc_dbcon.Execute ls_sql
                
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      If .TextMatrix(ln_cnt, 1) <> "" Then
                        ls_sql = "INSERT into IC_IssueNoteDetail(Compcode,BranchCode, TransCode,customcode, ItemCode, BalQty, Quantity, ItemRate, Amount,AvgRate,Remarks)"
                        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 8)) & "'," & Val(0 & .TextMatrix(ln_cnt, 4)) & "," & Val(0 & .TextMatrix(ln_cnt, 6)) & "," & Val(0 & .TextMatrix(ln_cnt, 5)) & "," & Val(0 & .TextMatrix(ln_cnt, 6)) * Val(0 & .TextMatrix(ln_cnt, 9)) & "," & Val(0 & .TextMatrix(ln_cnt, 9)) & " , '" & Trim(.TextMatrix(ln_cnt, 7)) & "' )"
                        gc_dbcon.Execute ls_sql
                     End If
                    Next
                 End With
                
                ls_sql = "delete from  IC_IssueNoteDetaillog where computername = '" & Gs_ComputerName & "' "
                gc_dbcon.Execute ls_sql
           
                 
     End Select
gc_dbcon.CommitTrans
If Mode <> "D" Then
   ls_opt = MsgBox("Print Issue Note ?.", vbYesNo)
   If ls_opt = vbYes Then Call PrintIssuenote
End If
If Mode = "A" Then
    txttransno = maxtranscode
     TxtRemarks = "Stock Issue to Showroom"
End If
InitializeGrid
Exit Sub
RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Public Sub ClearVal()
End Sub
Private Sub setprint()
End Sub
Private Sub PrintIssuenote()
On Error GoTo LocalErr

   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "issueNote.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Issue Note'"
        .SelectionFormula = "{PO_DemandNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.transcode} = '" & Trim(txttransno) & "' "
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub

Private Sub SetVal()
     Dim pr_dumy2 As New Recordset
     Dim ls_sql As String
     txtvaluedate = PR_ICIssue("Transdate")
     TxtRemarks = Trim(PR_ICIssue("Remarks") & "")
     txtGRNNo = Trim(PR_ICIssue("PCode") & "")
     
     ls_sql = "SELECT GRN.TransCode AS ComputerCode,  Vendors.Description  FROM         PO_POGRN GRN INNER JOIN         IC_Supplier Vendors ON GRN.Compcode = Vendors.Compcode AND GRN.AccountCode = Vendors.SupplierCode"
     ls_sql = ls_sql & " Where GRN.compcode = '" & Gs_compcode & "' and GRN.TransCode = '" & txtGRNNo & "'"
     If pr_dumy2.State = 1 Then pr_dumy2.Close
     pr_dumy2.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
     If Not pr_dumy2.EOF Then
     txtGrndesc = Trim(pr_dumy2("Description") & "")
     End If
     pr_dumy2.Close
     
     txtacode = Trim(PR_ICIssue("VCode") & "")
     Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
     txtacode1 = Trim(PR_ICIssue("ACode") & "")
     Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)
     
End Sub
Public Function ChkInputs() As Boolean
 Dim lb_opt As Boolean
    If Trim(TxtRemarks) = "" Then
      Call MsgBox("Enter Remarks !!!", vbCritical)
      ChkInputs = False
      If TxtRemarks.Enabled Then TxtRemarks.SetFocus
    ElseIf Trim(txtacode) = "" Then
      Call MsgBox("Enter/Select Verified Code !!!", vbCritical)
      ChkInputs = False
      txtacode.SetFocus
    ElseIf Trim(txtacode1) = "" Then
      Call MsgBox("Enter/Select Approved Code !!!", vbCritical)
      ChkInputs = False
      txtacode1.SetFocus
    
    ElseIf Not CheckValidQty Then
      ChkInputs = False
      
    ElseIf GrdGRN.TextMatrix(1, 1) = "" Then
      Call MsgBox("Enter Items in grid !!!", vbCritical)
      ChkInputs = False
      GrdGRN.SetFocus

    Else

        With GrdGRN
          For ln_cnt = 1 To .Rows - 1
           If .TextMatrix(ln_cnt, 1) <> "" Then
            If Val(.TextMatrix(ln_cnt, 5)) = 0 Then
                Call MsgBox("QTY Must be entered !!!", vbCritical)
                lb_opt = False
                GrdGRN.SetFocus
                Exit For
             Else
                lb_opt = True
             End If
            End If
          Next
        End With

       ChkInputs = lb_opt
    End If
End Function

Public Sub FrmRefresh()
    Pr_ICParty.Requery
    PR_ICIssue.Requery
    PR_IcItem.Requery
    PR_Branch.Requery
    PR_VchCntr.Requery
    PR_VchType.Requery
End Sub



Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Custom Code|<Item Name|<UOM|<Bal Qty|<Item Rate|<Qty|<Remarks|<Itemcode|<Avgrate"
        .ColWidth(1) = 1500
        .ColWidth(2) = 4000
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColAlignment(4) = 7
        .ColWidth(5) = 1100
        .ColAlignment(5) = 7
        .ColWidth(6) = 1100
        .ColAlignment(6) = 7
        .ColWidth(7) = 2500
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .Redraw = True
    End With
End Sub
Private Sub TotalAmount()
    txttotalamount = ""
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txttotalamount = Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 6))
        Next
    End With
    
End Sub
Function CheckValidQty() As Boolean

    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            If Val(.TextMatrix(ln_cnt, 4)) < Val(.TextMatrix(ln_cnt, 6)) Then
            Call MsgBox("Stock not available following item name " + .TextMatrix(ln_cnt, 2) & " Row # " + str(ln_cnt))
            CheckValidQty = False
            Exit For
            Else
            CheckValidQty = True
            End If
        Next
    End With
    
    
End Function

Private Sub LoadGRNTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String

ls_sql = "delete from  IC_IssueNoteDetaillog where computername = '" & Gs_ComputerName & "'"
gc_dbcon.Execute ls_sql

ls_sql = " SELECT IC_IssueNoteDetail.CustomCode,IC_IssueNoteDetail.ItemCode, IC_Item.Description, IC_IssueNoteDetail.BalQty, IC_IssueNoteDetail.Quantity, IC_IssueNoteDetail.ItemRate, IC_IssueNoteDetail.Amount, IC_ItemUM.Description AS UOM,IC_IssueNoteDetail.AvgRate,IC_IssueNoteDetail.Remarks "
ls_sql = ls_sql & " FROM IC_IssueNoteDetail INNER JOIN   IC_Item ON IC_IssueNoteDetail.Compcode = IC_Item.Compcode AND IC_IssueNoteDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where IC_IssueNoteDetail.Compcode = '" & Gs_compcode & "' and IC_IssueNoteDetail.Transcode = '" & txttransno & "' order by IC_IssueNoteDetail.rowid "

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("CustomCode") & "")
                .TextMatrix(.Row, 8) = Trim(Pr_LoadTrans("ItemCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("UOM") & "")
                .TextMatrix(.Row, 4) = Pr_LoadTrans("BalQty")
                .TextMatrix(.Row, 6) = Pr_LoadTrans("Quantity")
                .TextMatrix(.Row, 5) = Val(0 & Pr_LoadTrans("Itemrate"))
                .TextMatrix(.Row, 9) = Val(0 & Pr_LoadTrans("AvgRate"))
                .TextMatrix(.Row, 7) = Trim(Pr_LoadTrans("Remarks") & "")
           

 
           
                ls_sql = "INSERT into IC_IssueNoteDetaillog(Compcode,BranchCode, TransCode,customcode, ItemCode, Quantity, ItemRate, Amount,AvgRate,Remarks,Computername,rowid,EMode)"
                ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 8)) & "'," & Val(0 & .TextMatrix(.Row, 6)) & "," & Val(0 & .TextMatrix(.Row, 5)) & "," & Val(0 & .TextMatrix(.Row, 6)) * Val(0 & .TextMatrix(.Row, 9)) & "," & Val(0 & .TextMatrix(.Row, 9)) & " , '" & Trim(.TextMatrix(.Row, 7)) & "'  , '" & Trim(Gs_ComputerName) & "'," & .Row & " ,'" & Mode & "')"
                gc_dbcon.Execute ls_sql


                .Rows = .Rows + 1
                
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        
    Else
        Call SetErr("Transaction not found.!!!", vbCritical)
        
    End If
    Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub
Private Sub LoadLogTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String

ls_sql = " SELECT IC_IssueNoteDetail.CustomCode,IC_IssueNoteDetail.ItemCode, IC_Item.Description, IC_IssueNoteDetail.Quantity, IC_IssueNoteDetail.ItemRate, IC_IssueNoteDetail.Amount, IC_ItemUM.Description AS UOM,IC_IssueNoteDetail.AvgRate,IC_IssueNoteDetail.Remarks "
ls_sql = ls_sql & " FROM IC_IssueNoteDetailLog IC_IssueNoteDetail INNER JOIN   IC_Item ON IC_IssueNoteDetail.Compcode = IC_Item.Compcode AND IC_IssueNoteDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where IC_IssueNoteDetail.Compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'  order by IC_IssueNoteDetail.rowid "

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                 .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("CustomCode") & "")
                .TextMatrix(.Row, 8) = Trim(Pr_LoadTrans("ItemCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("UOM") & "")
                .TextMatrix(.Row, 6) = Pr_LoadTrans("Quantity")
                .TextMatrix(.Row, 5) = Val(0 & Pr_LoadTrans("Itemrate"))
                .TextMatrix(.Row, 9) = Val(0 & Pr_LoadTrans("AvgRate"))
                .TextMatrix(.Row, 7) = Trim(Pr_LoadTrans("Remarks") & "")
                .TextMatrix(.Row, 4) = 0

                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        
    Else
        Call SetErr("Transaction not found.!!!", vbCritical)
        
    End If
    Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub

Private Sub GrdGRN_DblClick()
    GrdGRN.SelectionMode = flexSelectionFree
End Sub
Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 And GrdGRN.Col = 1 Then ' F1 key pressed
  Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    
    Gs_SQL = "SELECT customCode,Description FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
    
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = Text1
    If GrdGRN.TextMatrix(GrdGRN.Row, 1) <> "" Then
        Call GrdGRN_KeyPress(13)
    End If
    
ElseIf KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With GrdGRN
           ls_sql = "delete from  IC_IssueNoteDetaillog where computername = '" & Gs_ComputerName & "' and rowid = " & .Row & ""
           gc_dbcon.Execute ls_sql
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
    End With
 ElseIf KeyCode = 46 Then 'Delete Key Pressed
   GrdGRN_KeyPress (KeyCode)
 End If

 
End Sub

Private Sub GrdGRN_KeyPress(KeyAscii As Integer)
'On Error GoTo ErrHandler
 Call GetKeysAdd(GrdGRN, KeyAscii)
Exit Sub
'ErrHandler:
'MsgBox ("An Error has Occured In The MSFlexgrid1_KeyPress() Procedure") & vbCr & "Report This Error To Latifjat@hotmail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.Number & vbCr & "Error Description : " & Err.Description, vbCritical, "FlexGrid Example"
End Sub
Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer)
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler

If KeyAscii = 13 Then 'if Enter Key then...
  Opt = ""
  With argFlexGrid
         .Row = .RowSel
    If .Col = 1 Then
           .CellBackColor = vbWindowBackground
           If .TextMatrix(.Row, 1) = "" Then
              Call GrdGRN_KeyDown(112, vbKeyShift)
           End If
          '.TextMatrix(.Row, 1) = DoPad(.TextMatrix(.Row, 1), 6)
          If SearchInGrid(GrdGRN, .TextMatrix(.Row, 1)) Then
             Call MsgBox("Item Code Already Exist in Grid ", vbCritical)
            .TextMatrix(.Row, 1) = ""
             Exit Sub
          End If

          If PR_IcItem.State = 1 Then PR_IcItem.Close
          PR_IcItem.Open " Select * From Ic_Item Where compcode = '" & Gs_compcode & "' and  CustomCode='" & Trim(.TextMatrix(.Row, 1)) & " '", gc_dbcon, adOpenStatic, adLockReadOnly
          
          If PR_IcItem.RecordCount <= 0 Then
              Call MsgBox(Gs_RecNFMsg, vbCritical)
             .TextMatrix(.Row, 1) = ""
             .Col = 1
          Else
             .TextMatrix(.Row, 0) = .Row
             .TextMatrix(.Row, 2) = Trim(PR_IcItem("Description") & "")
             .TextMatrix(.Row, 8) = Trim(PR_IcItem("Itemcode") & "")
             .TextMatrix(.Row, 5) = Val(PR_IcItem("SaleCost"))
             .TextMatrix(.Row, 9) = Val(PR_IcItem("AvgRate"))
             .TextMatrix(.Row, 4) = CheckBalQTY(.TextMatrix(.Row, 8), 1)
              txtitemname = .TextMatrix(.Row, 2)
             .Col = 6
             .CellBackColor = vbHighlight
              PR_UOM.Open "Select * From IC_ItemUM Where MCode='" & Trim(PR_IcItem("Mcode") & "") & " '", gc_dbcon, adOpenStatic, adLockReadOnly
              If PR_UOM.RecordCount > 0 Then
                .TextMatrix(.Row, 3) = Trim(PR_UOM("Description") & "")
              End If
              PR_UOM.Close
          End If
         PR_IcItem.Close
       ElseIf .Col = 2 Then
       ElseIf .Col = 3 Then
       ElseIf .Col = 6 Then
           .CellBackColor = vbWindowBackground
           If .TextMatrix(.Row, 6) = "" Then
             Call MsgBox("Enter Quantity!!!", vbCritical)
             Exit Sub
           End If
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
            
          End If
          .Col = 1
          .Row = .Row + 1
          .CellBackColor = vbHighlight
          .SetFocus
          
          
        If .RowSel > 9 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
            
   End If
   End With
 Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
If .Col = 1 Or .Col = 6 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .CellBackColor = vbWindowBackground
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
      Opt = ""
   End If
       ls_sql = "delete from  IC_IssueNoteDetaillog where computername = '" & Gs_ComputerName & "' and rowid = " & .Row & ""
       gc_dbcon.Execute ls_sql
           
       ls_sql = "INSERT into IC_IssueNoteDetaillog(Compcode,BranchCode, TransCode,customcode, ItemCode, Quantity, ItemRate, Amount,AvgRate,Remarks,Computername,rowid,EMode)"
       ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 8)) & "'," & Val(0 & .TextMatrix(.Row, 6)) & "," & Val(0 & .TextMatrix(.Row, 5)) & "," & Val(0 & .TextMatrix(.Row, 6)) * Val(0 & .TextMatrix(.Row, 9)) & "," & Val(0 & .TextMatrix(.Row, 9)) & " , '" & Trim(.TextMatrix(.Row, 7)) & "'  , '" & Trim(Gs_ComputerName) & "'," & .Row & " ,'" & Mode & "')"
       gc_dbcon.Execute ls_sql

     TotalAmount
End If
End With
End If

  If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      If .Col = 6 Then
         If .CellBackColor = vbHighlight Then
            .Text = "": .CellBackColor = vbWindowBackground
         End If
        .Text = .Text & Chr(KeyAscii)
         If Not IsNumeric(.Text) Then
            Call MsgBox("Enter numeric entry !!!", vbCritical)
            .Text = ""
         End If
         If Val(.TextMatrix(.Row, 6)) > Val(.TextMatrix(.Row, 4)) Then
           Call MsgBox("Qty Less then Bal QTY !!!", vbCritical)
          .Text = ""
         End If
     
'         If Val(.Text) > Val(.TextMatrix(.Row, 9)) - Val(.TextMatrix(.Row, 10)) Then
'           Call MsgBox("Qty Less then Job QTY !!!", vbCritical)
'          .Text = ""
'         End If
'         .TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 4)) * .TextMatrix(.Row, 5)
           ls_sql = "delete from  IC_IssueNoteDetaillog where computername = '" & Gs_ComputerName & "' and rowid = " & .Row & ""
           gc_dbcon.Execute ls_sql
           
           ls_sql = "INSERT into IC_IssueNoteDetaillog(Compcode,BranchCode, TransCode,customcode, ItemCode, Quantity, ItemRate, Amount,AvgRate,Remarks,Computername,rowid,EMode)"
           ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 8)) & "'," & Val(0 & .TextMatrix(.Row, 6)) & "," & Val(0 & .TextMatrix(.Row, 5)) & "," & Val(0 & .TextMatrix(.Row, 6)) * Val(0 & .TextMatrix(.Row, 9)) & "," & Val(0 & .TextMatrix(.Row, 9)) & " , '" & Trim(.TextMatrix(.Row, 7)) & "'  , '" & Trim(Gs_ComputerName) & "'," & .Row & ",'" & Mode & "' )"
           gc_dbcon.Execute ls_sql


          TotalAmount
    ElseIf .Col = 1 Or .Col = 7 Then
        If .CellBackColor = vbHighlight Then
            .Text = "": .CellBackColor = vbWindowBackground
        End If
      .Text = .Text & Chr(KeyAscii)
    End If
        
    End With
  End If
End Sub
Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TxtRemarks.SetFocus
    End If
End Sub

Public Sub SetFrmEnv(ls_mode As String)
  
End Sub

