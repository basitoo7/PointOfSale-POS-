VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInventoryWaste 
   Caption         =   "Inventory Waste"
   ClientHeight    =   7320
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
   Icon            =   "FrmInventoryWaste.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
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
      Begin VB.ComboBox txtadjin 
         Height          =   330
         ItemData        =   "FrmInventoryWaste.frx":030A
         Left            =   10140
         List            =   "FrmInventoryWaste.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   150
         Width           =   1935
      End
      Begin VB.ComboBox txtinventorytype 
         Height          =   330
         ItemData        =   "FrmInventoryWaste.frx":032A
         Left            =   6645
         List            =   "FrmInventoryWaste.frx":0334
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   165
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1500
         Picture         =   "FrmInventoryWaste.frx":0355
         Style           =   1  'Graphical
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1140
         Width           =   4035
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   7470
         Picture         =   "FrmInventoryWaste.frx":04C7
         Style           =   1  'Graphical
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2160
         Picture         =   "FrmInventoryWaste.frx":0639
         Style           =   1  'Graphical
         TabIndex        =   12
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
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1110
         Visible         =   0   'False
         Width           =   195
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   3915
         TabIndex        =   8
         Top             =   165
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16384001
         CurrentDate     =   37580
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3870
         MaxLength       =   50
         TabIndex        =   7
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
      Begin VB.Label Label1 
         Caption         =   "Waste IN :"
         Height          =   255
         Left            =   9285
         TabIndex        =   27
         ToolTipText     =   "Enter Value Date"
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label9 
         Caption         =   "Adj Type :"
         Height          =   255
         Left            =   5895
         TabIndex        =   25
         ToolTipText     =   "Enter Value Date"
         Top             =   195
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Verified By :"
         Height          =   255
         Left            =   75
         TabIndex        =   23
         Top             =   1185
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Approved By :"
         Height          =   255
         Left            =   5955
         TabIndex        =   22
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Waste #  :"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   195
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   555
         Width           =   780
      End
      Begin VB.Label label2 
         Caption         =   "Waste Date :"
         Height          =   255
         Left            =   2910
         TabIndex        =   4
         ToolTipText     =   "Enter Value Date"
         Top             =   180
         Width           =   1065
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
               Picture         =   "FrmInventoryWaste.frx":07AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryWaste.frx":0BFF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryWaste.frx":1053
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryWaste.frx":14A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryWaste.frx":18FB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryWaste.frx":1D4F
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryWaste.frx":24A3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5160
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   10500
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   4740
         Width           =   1590
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   4470
         Left            =   150
         TabIndex        =   15
         Top             =   195
         Width           =   11970
         _ExtentX        =   21114
         _ExtentY        =   7885
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
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   4755
         Width           =   5310
      End
      Begin VB.TextBox txtnoofitems 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label4 
         Caption         =   " Total :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9990
         TabIndex        =   29
         Top             =   4770
         Width           =   1020
      End
      Begin VB.Label Label11 
         Caption         =   " Total :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6345
         TabIndex        =   6
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
Attribute VB_Name = "FrmInventoryWaste"
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
Dim ln_Adj As Integer
Dim ls_transtype As String

Dim pr_dumy As New Recordset

Dim PR_UOM As New Recordset

Dim PR_ICIssue As New Recordset
Dim PR_IcItem As New Recordset
Dim PR_Branch As New Recordset
Private Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from IC_InventoryWastMaster where compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
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
End Sub

Private Sub New_Record_Click()
       txttransno.Enabled = False
       Command1.Enabled = False
       InitializeGrid
       txttransno = maxtranscode
       Mode = DentMode(Mode, 1, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
        TxtRemarks = "Adjustment of Inventory"
       TxtRemarks.SetFocus
       CheckLogTrans
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
            GrdGRN.Col = 1
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
    Gs_SQL = "Select TransCode, Transdate from IC_InventoryWastMaster "
    Gs_FindFld = "TransCode"
    Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "'  and glstatus = 0"
    Gs_OrderBy = "Order by TransCode"

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
  SetToolBar(2) = chkRights("INVISSUE02")
  SetToolBar(3) = chkRights("INVISSUE03")
  SetToolBar(4) = chkRights("INVISSUE04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  txtinventorytype.ListIndex = 0

  txtvaluedate.Value = Date
 
  InitializeGrid

  
End Sub

Private Sub txtinventorytype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtadjin.SetFocus
End Sub
Private Sub txtadjin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then TxtRemarks.SetFocus
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
         PR_ICIssue.Open "select * from IC_InventoryWastMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0 ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
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
                   LoadGRNTrans
                   If Mode <> "D" Then
                      txttransno.SetFocus
                   End If
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
       TxtRemarks = "Waste of Inventory"
       txtadjin.SetFocus
       'CheckLogTrans
    End If
End Sub


Public Sub SaveValues()
'On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String

gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
              gc_dbcon.Execute "DELETE FROM IC_InventoryWastMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              gc_dbcon.Execute "DELETE FROM IC_InventoryWastDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              
              
           Case Else
                If Mode = "E" Then
                    gc_dbcon.Execute "DELETE FROM IC_InventoryWastMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                    gc_dbcon.Execute "DELETE FROM IC_InventoryWastDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                End If
                If Mode = "A" Then
                    txttransno = maxtranscode
                    Me.Refresh
                End If
                    
                      ls_sql = "INSERT into IC_InventoryWastMaster( Compcode,branchcode, TransCode,   TransDate, Remarks,vcode,acode,userid,adddate,addtime,Adjsiteid)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & RepApp(TxtRemarks) & "','" & txtacode & "','" & txtacode1 & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "'," & txtadjin.ListIndex + 1 & " )"
                      gc_dbcon.Execute ls_sql
                
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                       
                    If .TextMatrix(ln_cnt, 7) = "Inv Out" Then
                        ln_Adj = 1
                    Else
                        ln_Adj = 0
                    End If
                       
                      If .TextMatrix(ln_cnt, 1) <> "" Then
                        ls_sql = "INSERT into IC_InventoryWastDetail(Compcode,BranchCode, TransCode,customcode, ItemCode,Stockinhand, Quantity,InvType, ItemRate, Amount,AvgRate,rowid)"
                        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 12)) & "'," & Val(0 & .TextMatrix(ln_cnt, 4)) & "," & Val(0 & .TextMatrix(ln_cnt, 6)) & "," & ln_Adj & "," & Val(0 & .TextMatrix(ln_cnt, 8)) & "," & Val(0 & .TextMatrix(ln_cnt, 9)) & "," & Val(0 & .TextMatrix(ln_cnt, 10)) & "," & Val(ln_cnt) & "  )"
                        gc_dbcon.Execute ls_sql
                     End If
                    Next
                 End With
                'ls_sql = "Delete from  IC_InventoryWastDetailLog where computername ='" & Gs_ComputerName & "'"
                'gc_dbcon.Execute ls_sql
     End Select
gc_dbcon.CommitTrans
If Mode <> "D" Then
   ls_opt = MsgBox("Print Waste Note ?.", vbYesNo)
   If ls_opt = vbYes Then Call PrintIssuenote
End If
If Mode = "A" Then
     txttransno = maxtranscode
     TxtRemarks = "Waste of Inventory"
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
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "WasteInventoryNote.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Waste Inventory Note'"
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
     
     txtvaluedate = PR_ICIssue("Transdate")
     TxtRemarks = Trim(PR_ICIssue("Remarks") & "")
     txtacode = Trim(PR_ICIssue("VCode") & "")
     Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
     txtacode1 = Trim(PR_ICIssue("ACode") & "")
     Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)
     txtadjin.ListIndex = Val(0 & PR_ICIssue("Adjsiteid")) - 1
   
End Sub
Public Function ChkInputs() As Boolean
 Dim lb_opt As Boolean
    If Trim(TxtRemarks) = "" Then
      Call MsgBox("Enter Remarks !!!", vbCritical)
      ChkInputs = False
      If TxtRemarks.Enabled Then TxtRemarks.SetFocus
    ElseIf Trim(txtadjin.Text) = "" Then
      Call MsgBox("Select Adjustment In !!!", vbCritical)
      ChkInputs = False
      If txtadjin.Enabled Then txtadjin.SetFocus
    ElseIf Trim(txtacode) = "" Then
      Call MsgBox("Enter/Select Verified Code !!!", vbCritical)
      ChkInputs = False
      txtacode.SetFocus
    ElseIf Trim(txtacode1) = "" Then
      Call MsgBox("Enter/Select Approved Code !!!", vbCritical)
      ChkInputs = False
      txtacode1.SetFocus
    ElseIf GrdGRN.TextMatrix(1, 1) = "" Then
      Call MsgBox("Enter Items in grid !!!", vbCritical)
      ChkInputs = False
      GrdGRN.SetFocus

    Else
      ChkInputs = True
    
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
        .FormatString = "Sr# |<Custom Code|<Item Name|<UOM|<Stock In Hand|<Stock on Shelf|<Quantity|<Adj Type|<Rate|<Total|<ItemRate|<Remarks|<Itemcode"
        .ColWidth(1) = 1500
        .ColWidth(2) = 3800
        .ColWidth(3) = 0
        .ColWidth(4) = 1200
        .ColAlignment(4) = 7
        .ColWidth(5) = 0
        .ColAlignment(5) = 7
        .ColWidth(6) = 1300
        .ColAlignment(6) = 7
        .ColWidth(7) = 1100
        .ColWidth(8) = 1100
        .ColWidth(9) = 1200
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        
        .Redraw = True
    End With
End Sub

Private Sub TotalAmount()
    txttotalamount = ""
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txttotalamount = Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 9))
        Next
    End With
End Sub
Private Sub LoadGRNTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String


ls_sql = " SELECT IC_InventoryWastDetail.CustomCode,IC_InventoryWastDetail.ItemCode, IC_Item.Description,IC_InventoryWastDetail.invtype,IC_InventoryWastDetail.Stockinhand,IC_InventoryWastDetail.Quantity, IC_InventoryWastDetail.ItemRate, IC_InventoryWastDetail.Amount, IC_ItemUM.Description AS UOM,IC_InventoryWastDetail.AvgRate "
ls_sql = ls_sql & " FROM IC_InventoryWastDetail INNER JOIN   IC_Item ON IC_InventoryWastDetail.Compcode = IC_Item.Compcode AND IC_InventoryWastDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where IC_InventoryWastDetail.Compcode = '" & Gs_compcode & "' and IC_InventoryWastDetail.Transcode = '" & txttransno & "' order by IC_InventoryWastDetail.rowid"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                 .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("CustomCode") & "")
                .TextMatrix(.Row, 12) = Trim(Pr_LoadTrans("ItemCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("UOM") & "")
                .TextMatrix(.Row, 4) = Val(Pr_LoadTrans("Stockinhand"))
                .TextMatrix(.Row, 6) = Val(Pr_LoadTrans("Quantity"))
               
                
                If Val(Pr_LoadTrans("Invtype")) = 1 Then
                    .TextMatrix(.Row, 7) = "Inv Out"
                Else
                   .TextMatrix(.Row, 7) = "Inv In"
                End If
                
                
                .TextMatrix(.Row, 8) = Val(0 & Pr_LoadTrans("Itemrate"))
                .TextMatrix(.Row, 9) = Val(0 & Pr_LoadTrans("amount"))
                .TextMatrix(.Row, 10) = Val(0 & Pr_LoadTrans("AvgRate"))
                
                .Rows = .Rows + 1
                
'                ls_sql = "INSERT into IC_InventoryWastDetailLog(Compcode,BranchCode, TransCode,customcode, ItemCode, Stockinhand,Quantity,stockonshelf,invtype, ItemRate, Amount,AvgRate,Remarks,ComputerName,rowid,EMode)"
'                ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 12)) & "'," & Val(0 & .TextMatrix(.Row, 4)) & "," & Val(0 & .TextMatrix(.Row, 5)) & "," & Val(0 & .TextMatrix(.Row, 6)) & "," & Val(Pr_LoadTrans("Invtype")) & "," & Val(0 & .TextMatrix(.Row, 8)) & "," & Val(0 & .TextMatrix(.Row, 9)) & "," & Val(0 & .TextMatrix(.Row, 10)) & " , '" & Trim(.TextMatrix(.Row, 11)) & "','" & Gs_ComputerName & "'," & .Row & " ,'" & Mode & "')"
'                gc_dbcon.Execute ls_sql
                
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

ls_sql = " SELECT IC_InventoryWastDetail.CustomCode,IC_InventoryWastDetail.ItemCode, IC_Item.Description, IC_InventoryWastDetail.stockinhand,IC_InventoryWastDetail.stockonshelf,IC_InventoryWastDetail.invtype,IC_InventoryWastDetail.Quantity, IC_InventoryWastDetail.ItemRate, IC_InventoryWastDetail.Amount, IC_ItemUM.Description AS UOM,IC_InventoryWastDetail.AvgRate,IC_InventoryWastDetail.Remarks "
ls_sql = ls_sql & " FROM IC_InventoryWastDetaillog IC_InventoryWastDetail INNER JOIN   IC_Item ON IC_InventoryWastDetail.Compcode = IC_Item.Compcode AND IC_InventoryWastDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where IC_InventoryWastDetail.Compcode = '" & Gs_compcode & "' and computername ='" & Gs_ComputerName & "' order by IC_InventoryWastDetail.rowid"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                 .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("CustomCode") & "")
                .TextMatrix(.Row, 12) = Trim(Pr_LoadTrans("ItemCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("UOM") & "")
                .TextMatrix(.Row, 4) = Pr_LoadTrans("Stockinhand")
                .TextMatrix(.Row, 5) = Pr_LoadTrans("Stockonshelf")
                .TextMatrix(.Row, 6) = Pr_LoadTrans("Quantity")
                 
                 If Val(Pr_LoadTrans("Invtype")) = 0 Then
                    .TextMatrix(.Row, 7) = "Inv In"
                 Else
                    .TextMatrix(.Row, 7) = "Inv Out"
                 End If
                
                .TextMatrix(.Row, 8) = Val(0 & Pr_LoadTrans("Itemrate"))
                .TextMatrix(.Row, 9) = Val(0 & Pr_LoadTrans("amount"))
                .TextMatrix(.Row, 10) = Val(0 & Pr_LoadTrans("AvgRate"))
                .TextMatrix(.Row, 11) = Trim(Pr_LoadTrans("Remarks") & "")
                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
    End If
    Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub
Private Sub CheckLogTrans()
Dim pr_dumyLog As New Recordset
Dim res
pr_dumyLog.Open "select * from IC_InventoryWastDetaillog  where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyLog.EOF Then
    If pr_dumyLog("Emode") = "E" Then
        txttransno = pr_dumyLog("Transcode")
        res = MsgBox(txttransno & " # you have opened in edit mode not save Do you want to open now", vbYesNo + vbExclamation)
        If res = vbYes Then
        Mode = DentMode(Mode, 2, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
        If PR_ICIssue.State = 1 Then PR_ICIssue.Close
        PR_ICIssue.Open "select * from IC_InventoryWastMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and glstatus = 0 ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If Not PR_ICIssue.EOF Then
        Call SetVal
        End If
        PR_ICIssue.Close
        LoadLogTrans
        Else
        
        ls_sql = "delete from IC_InventoryWastDetaillog  where computername = '" & Gs_ComputerName & "' "
        gc_dbcon.Execute ls_sql
           
        End If
    Else
        LoadLogTrans
    End If
End If
pr_dumyLog.Close
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
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            ls_sql = "delete from  IC_InventoryWastDetailLog  where computername = '" & Gs_ComputerName & "' and rowid = " & .Row & ""
            gc_dbcon.Execute ls_sql
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
          'If SearchInGrid(GrdGRN, .TextMatrix(.Row, 1)) Then
           '  Call MsgBox("Record Already Exist in Grid", vbCritical)
           ' .TextMatrix(.Row, 1) = ""
            ' Exit Sub
          'End If

          If PR_IcItem.State = 1 Then PR_IcItem.Close
          PR_IcItem.Open " Select * From Ic_Item Where compcode = '" & Gs_compcode & "' and  CustomCode='" & Trim(.TextMatrix(.Row, 1)) & " '", gc_dbcon, adOpenStatic, adLockReadOnly
          
          If PR_IcItem.RecordCount <= 0 Then
              Call MsgBox(Gs_RecNFMsg, vbCritical)
             .TextMatrix(.Row, 1) = ""
          Else
             .TextMatrix(.Row, 0) = .Row
             .TextMatrix(.Row, 2) = Trim(PR_IcItem("Description") & "")
             .TextMatrix(.Row, 12) = Trim(PR_IcItem("Itemcode") & "")
             .TextMatrix(.Row, 4) = CheckBalQTY(.TextMatrix(.Row, 12), txtadjin.ListIndex + 1)
             .TextMatrix(.Row, 8) = PR_IcItem("Purchasecost")
             .TextMatrix(.Row, 10) = PR_IcItem("Purchasecost")
             .TextMatrix(.Row, 9) = 0
             .TextMatrix(.Row, 10) = 0
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

          If PR_IcItem.State = 1 Then PR_IcItem.Close
          PR_IcItem.Open " Select * From Ic_Item Where compcode = '" & Gs_compcode & "' and  ItemCode='" & Trim(.TextMatrix(.Row, 12)) & " '", gc_dbcon, adOpenStatic, adLockReadOnly
          If Not PR_IcItem.EOF Then
           .TextMatrix(.Row, 10) = PR_IcItem("AvgRate")
           .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 6)) * .TextMatrix(.Row, 8)
          End If
             
          PR_IcItem.Close
          Call TotalAmount
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
If .Col = 6 Or .Col = 1 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .CellBackColor = vbWindowBackground
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
      Opt = ""
   End If
   

If Val(.TextMatrix(.Row, 6)) > 0 Then
            'If Val(.TextMatrix(.Row, 4)) > Val(.TextMatrix(.Row, 5)) Then
            '.TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 4)) - Val(.TextMatrix(.Row, 5))
            '.TextMatrix(.Row, 7) = "Inv Out"
            'ln_Adj = 1
            'Else
          '  .TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 5)) - Val(.TextMatrix(.Row, 4))
           .TextMatrix(.Row, 7) = "Inv Out"
            ln_Adj = 0
           ' End If
            
         .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 6)) * Val(.TextMatrix(.Row, 8))
Else
'.TextMatrix(.Row, 6) = ""
'.TextMatrix(.Row, 7) = ""
.TextMatrix(.Row, 9) = ""
End If

'ls_sql = "Delete from  IC_InventoryWastDetailLog where computername ='" & Gs_ComputerName & "' and rowid = " & .Row & ""
'gc_dbcon.Execute ls_sql

'ls_sql = "INSERT into IC_InventoryWastDetailLog(Compcode,BranchCode, TransCode,customcode, ItemCode,stockinhand,stockonshelf,invtype, Quantity, ItemRate, Amount,AvgRate,Remarks,ComputerName,rowid,EMode)"
'ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 12)) & "'," & Val(0 & .TextMatrix(.Row, 4)) & "," & Val(0 & .TextMatrix(.Row, 5)) & "," & ln_Adj & "," & Val(0 & .TextMatrix(.Row, 6)) & "," & Val(0 & .TextMatrix(.Row, 8)) & "," & Val(0 & .TextMatrix(.Row, 9)) & "," & Val(0 & .TextMatrix(.Row, 10)) & " , '" & Trim(.TextMatrix(.Row, 11)) & "','" & Gs_ComputerName & "'," & .Row & ",'" & Mode & "' )"
'gc_dbcon.Execute ls_sql
             
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
'         If Val(.Text) > Val(.TextMatrix(.Row, 11)) Then
'           Call MsgBox("Qty Less then Closing QTY !!!", vbCritical)
'          .Text = ""
'         End If
     
'         If Val(.Text) > Val(.TextMatrix(.Row, 9)) - Val(.TextMatrix(.Row, 10)) Then
'           Call MsgBox("Qty Less then Job QTY !!!", vbCritical)
'          .Text = ""
'         End If
        
'        If Val(.TextMatrix(.Row, 5)) = 0 Then
'        .TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 4))
'        .TextMatrix(.Row, 7) = "Inv Out"
'         ln_Adj = 1
'         End If
'
        If Val(.TextMatrix(.Row, 6)) > 0 Then
'            If Val(.TextMatrix(.Row, 4)) > Val(.TextMatrix(.Row, 5)) Then
'            .TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 4)) - Val(.TextMatrix(.Row, 5))
'            .TextMatrix(.Row, 7) = "Inv Out"
'            ln_Adj = 1
'            Else
'            .TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 5)) - Val(.TextMatrix(.Row, 4))
            .TextMatrix(.Row, 7) = "Inv Out"
            ln_Adj = 0
        End If
      
          .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 6)) * .TextMatrix(.Row, 8)
          TotalAmount
 '         ls_sql = "Delete from  IC_InventoryWastDetailLog where computername ='" & Gs_ComputerName & "' and rowid = " & .Row & ""
 '         gc_dbcon.Execute ls_sql
          
  '         ls_sql = "INSERT into IC_InventoryWastDetailLog(Compcode,BranchCode, TransCode,customcode, ItemCode,stockinhand,stockonshelf,invtype, Quantity, ItemRate, Amount,AvgRate,Remarks,ComputerName,rowid,EMode)"
   '        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 12)) & "'," & Val(0 & .TextMatrix(.Row, 4)) & "," & Val(0 & .TextMatrix(.Row, 5)) & "," & ln_Adj & "," & Val(0 & .TextMatrix(.Row, 6)) & "," & Val(0 & .TextMatrix(.Row, 8)) & "," & Val(0 & .TextMatrix(.Row, 9)) & "," & Val(0 & .TextMatrix(.Row, 10)) & " , '" & Trim(.TextMatrix(.Row, 11)) & "','" & Gs_ComputerName & "'," & .Row & ",'" & Mode & "' )"
    '       gc_dbcon.Execute ls_sql
    ElseIf .Col = 1 Or .Col = 8 Then
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
        txtadjin.SetFocus
    End If
End Sub

Public Sub SetFrmEnv(ls_mode As String)
  
End Sub

