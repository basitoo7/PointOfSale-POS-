VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOVendorPayments 
   Caption         =   "Payments to Vendors By Cash"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendorPayments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11880
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
      Height          =   5475
      Left            =   0
      TabIndex        =   3
      Top             =   570
      Width           =   11910
      Begin VB.TextBox txtpayableamount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         MaxLength       =   64
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4185
         Width           =   1410
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2895
         TabIndex        =   32
         Text            =   "Text2"
         Top             =   240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox txtrecamount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   10320
         MaxLength       =   64
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4185
         Width           =   1410
      End
      Begin VB.TextBox txtnetamount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4950
         MaxLength       =   64
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4185
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2505
         Picture         =   "frmVendorPayments.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   -15
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   195
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   10080
         TabIndex        =   0
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63045633
         CurrentDate     =   37580
      End
      Begin VB.TextBox txtVendorDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2430
         MaxLength       =   64
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   9375
      End
      Begin VB.TextBox txtVendorCode 
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
         MaxLength       =   6
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   660
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   2085
         Picture         =   "frmVendorPayments.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   315
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
         Height          =   3300
         Left            =   75
         TabIndex        =   12
         Top             =   825
         Width           =   11760
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
            Height          =   3030
            Left            =   75
            TabIndex        =   13
            Top             =   195
            Width           =   11580
            _ExtentX        =   20426
            _ExtentY        =   5345
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
      Begin VB.Frame Frame3 
         Height          =   1035
         Left            =   105
         TabIndex        =   18
         Top             =   4440
         Width           =   11805
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
            Left            =   4125
            MaxLength       =   200
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   600
            Width           =   7590
         End
         Begin VB.TextBox txtbankdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5130
            MaxLength       =   64
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   225
            Width           =   4035
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
            Left            =   4125
            MaxLength       =   4
            TabIndex        =   24
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   225
            Width           =   660
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   4800
            Picture         =   "frmVendorPayments.frx":05EE
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   225
            Width           =   315
         End
         Begin VB.ComboBox txtpaymenttype 
            Height          =   330
            ItemData        =   "frmVendorPayments.frx":0760
            Left            =   1245
            List            =   "frmVendorPayments.frx":076D
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   180
            Width           =   1935
         End
         Begin VB.TextBox txtinstrumentNo 
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
            Left            =   1245
            MaxLength       =   200
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   600
            Width           =   1905
         End
         Begin VB.TextBox txtvchrtype 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   9960
            MaxLength       =   64
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   225
            Width           =   675
         End
         Begin VB.TextBox txtvchrno 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10665
            MaxLength       =   64
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Remarks :"
            Height          =   255
            Left            =   3120
            TabIndex        =   31
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bank Code :"
            Height          =   210
            Index           =   2
            Left            =   3210
            TabIndex        =   30
            Top             =   255
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Payment Type :"
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   210
            Width           =   1110
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Instrument # :"
            Height          =   255
            Left            =   225
            TabIndex        =   28
            Top             =   615
            Width           =   990
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Voucher :"
            Height          =   255
            Left            =   9240
            TabIndex        =   27
            Top             =   255
            Width           =   720
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Balance Payable :"
         Height          =   210
         Index           =   6
         Left            =   6510
         TabIndex        =   34
         Top             =   4215
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Paid Amount :"
         Height          =   210
         Index           =   5
         Left            =   9330
         TabIndex        =   17
         Top             =   4215
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount :"
         Height          =   210
         Index           =   4
         Left            =   3885
         TabIndex        =   16
         Top             =   4215
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Trans Code :"
         Height          =   255
         Left            =   75
         TabIndex        =   11
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Code :"
         Height          =   210
         Index           =   1
         Left            =   300
         TabIndex        =   7
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Value Date :"
         Height          =   255
         Index           =   0
         Left            =   9150
         TabIndex        =   4
         ToolTipText     =   "Enter Value Date"
         Top             =   165
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
               Picture         =   "frmVendorPayments.frx":0784
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPayments.frx":0BD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPayments.frx":102C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPayments.frx":1480
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPayments.frx":18D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPayments.frx":1D28
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendorPayments.frx":247C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu Filemenu 
      Caption         =   "File"
      Begin VB.Menu NewRecord 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu EditRecord 
         Caption         =   "&Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu DeleteRecord 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu SaveRecord 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Editmenu 
      Caption         =   "Edit"
   End
End
Attribute VB_Name = "frmPOVendorPayments"
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
Private Sub TotalAmount()
    Dim ln_cnt As Integer
      txtNetAmount = ""
      txtrecamount = ""
      txtpayableamount = ""
      
      
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txtNetAmount = Val(txtNetAmount) + Val(.TextMatrix(ln_cnt, 6))
            txtpayableamount = Val(txtpayableamount) + Val(.TextMatrix(ln_cnt, 7))
            txtrecamount = Val(txtrecamount) + Val(.TextMatrix(ln_cnt, 8))
        Next
    End With
End Sub

Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<GRN Code|<Amount|<GST Amount|<SED Amount|<Dis Amount |<Total Amount|<Balance Payable|<Paid Amount|<Remarks"
        .ColWidth(1) = 1150
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColAlignment(4) = 7
        .ColWidth(5) = 1400
        .ColAlignment(5) = 7
        .ColWidth(6) = 1400
        .ColWidth(7) = 1400
        .ColWidth(8) = 1400
        .ColWidth(9) = 4000
        .Redraw = True
    End With
End Sub
Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer)
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler

If KeyAscii = 13 Then 'if Enter Key then...
  Opt = ""
  With argFlexGrid
        .SelectionMode = flexSelectionByRow
        Row = .RowSel
       If .Col = 1 Then
       .Col = .Col + 6
       .CellBackColor = vbHighlight

       ElseIf .Col = 2 Then
       ElseIf .Col = 3 Then
       ElseIf .Col = 6 Then
           If .TextMatrix(Row, 6) = "" Then
             Call MsgBox("Enter Amount!!!", vbCritical)
             Exit Sub
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
    If .Col = 8 Then
      .CellBackColor = vbWindowBackground
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
      Opt = ""
    End If
   End If
End With
End If

  If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      If .Col = 8 Then
         If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
         End If
         .Text = .Text & Chr(KeyAscii)
          If Not IsNumeric(.Text) Then
          .Text = ""
           Call MsgBox("Enter Numeric entry !!!", vbCritical)
          End If
       If Val(.TextMatrix(.Row, 8)) > Val(.TextMatrix(.Row, 7)) Then
       Call MsgBox("Paid amount not greater then balance payable !!!", vbCritical)
        .TextMatrix(.Row, 8) = Val(.TextMatrix(.Row, 7))
       End If
          
       TotalAmount
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

Private Sub NewRecord_Click()
Mode = DentMode(Mode, 1, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
Command1.Enabled = False
InitializeGrid
txttransno = maxtranscode
txttransno.Enabled = False

End Sub
Private Sub saverecord_Click()
Mode = DentMode(Mode, 4, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
If txtVendorCode.Enabled Then txtVendorCode.SetFocus
Command1.Enabled = True
txttransno.Enabled = True
End Sub
Private Sub editrecord_Click()
Mode = DentMode(Mode, 2, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
If txtVendorCode.Enabled Then txtVendorCode.SetFocus
Command1.Enabled = True
txttransno.Enabled = True
txttransno.SetFocus
End Sub
Private Sub Deleterecord_Click()
Mode = DentMode(Mode, 3, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
If txtVendorCode.Enabled Then txtVendorCode.SetFocus
Command1.Enabled = True
txttransno.Enabled = True
txttransno.SetFocus
End Sub

Private Sub GrdGRN_Click()
GrdGRN.SelectionMode = flexSelectionFree
With GrdGRN
txtitemname = .TextMatrix(.Row, 2)
End With
GrdGRN.CellBackColor = vbHighlight
End Sub

Private Sub GrdGRN_DblClick()
    GrdGRN.SelectionMode = flexSelectionFree
End Sub

Private Sub GrdGRN_EnterCell()
GrdGRN.CellBackColor = vbHighlight
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 113 Then  ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2



    Gs_SQL = " SELECT PO_POOrderNote.TransDate, IC_Item.Description, IC_ItemUM.Description AS UOM, PO_POOrderDetail.Quantity, PO_POOrderDetail.BonusQty,"
    Gs_SQL = Gs_SQL & " PO_POOrderDetail.BonusAmount, PO_POOrderDetail.Rate, PO_POOrderDetail.Amount, PO_POOrderDetail.GSTAmount, PO_POOrderDetail.SEDAmount,"
    Gs_SQL = Gs_SQL & " PO_POOrderDetail.DiscAmount FROM PO_POGRN PO_POOrderNote INNER JOIN PO_POGRNDetail PO_POOrderDetail ON PO_POOrderNote.Compcode = PO_POOrderDetail.Compcode AND"
    Gs_SQL = Gs_SQL & " PO_POOrderNote.BranchCode = PO_POOrderDetail.BranchCode AND PO_POOrderNote.TransCode = PO_POOrderDetail.TransCode INNER JOIN"
    Gs_SQL = Gs_SQL & " IC_Item IC_Item ON PO_POOrderDetail.Compcode = IC_Item.Compcode AND PO_POOrderDetail.ItemCode = IC_Item.ItemCode INNER JOIN"
    Gs_SQL = Gs_SQL & " IC_ItemUM IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
    Gs_FindFld = "IC_Item.Description"
    Gs_OrderBy = "ORDER BY PO_POOrderNote.TransDate"

    Gs_OtherPara = " where PO_POOrderNote.compcode = '" & Gs_compcode & "' and PO_POOrderNote.TransCode = '" & GrdGRN.TextMatrix(GrdGRN.Row, 1) & "'   "

    MyLookupMultifields.Caption = "Order Detail"
    MyLookupMultifields.Show 1 '
 ElseIf KeyCode = vbKeyDelete Then  'Delete Key Pressed
            With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
            TotalAmount
           End With
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


Private Sub txtbankcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtbankcode <> "" Then
    txtbankcode = DoPad(txtbankcode, txtbankcode.MaxLength)
    
        ls_sql = "Select bankcode,bankname as Description from SysBanks where bankcode = '" & txtbankcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Bank code not found", vbCritical)
                'Cancel = True
            Else
                txtbankdesc = pr_dumy("description")
                If txtinstrumentNo.Enabled Then txtinstrumentNo.SetFocus
            End If
         pr_dumy.Close
    
End If

End Sub

Private Sub txtinstrumentNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then TxtRemarks.SetFocus
End Sub
Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttransno
    Set PO_DESC = Text1
    Gs_SQL = "Select TransCode, Transdate from PO_PayableMaster"
    Gs_FindFld = "TransCode"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' and GLstatus = 0"
    Gs_OrderBy = "Order by TransCode"

    MyLookupOLDB.Caption = "Vendor Payments"
    MyLookupOLDB.Show 1
    
    If txttransno <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtVendorCode
    Set PO_DESC = txtVendordesc
    Gs_SQL = "Select  SupplierCode 'Code' ,Description from Ic_Supplier"
    Gs_FindFld = "Description"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Description,SupplierCode"
    
    MyLookupOLDB.Caption = "Vendors"
    MyLookupOLDB.Show 1
    
    If Len(txtVendorCode) > 0 Then txtVendorCode_KeyDown vbKeyReturn, vbKeyShift
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
       Mode = DentMode(Mode, Button.Index, PR_ICIssue, Me, txtvaluedate, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
       If Mode = "A" Then
        txttransno = maxtranscode
        InitializeGrid
        txttransno.Enabled = False
        Command1.Enabled = False
        txtVendorCode.SetFocus
        
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
Dim ls_transtype As String


gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
                        gc_dbcon.Execute "DELETE FROM PO_PayableMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                        gc_dbcon.Execute "DELETE FROM PO_PayableDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              
                         With GrdGRN
                                For ln_cnt = 1 To .Rows - 1
                                      If Val(.TextMatrix(ln_cnt, 8)) > 0 Then
                                        ls_sql = "Update  PO_POGRN set RecAmount = " & CheckBalAmount(Trim(.TextMatrix(ln_cnt, 1))) & "  WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(.TextMatrix(ln_cnt, 1)) & "'"
                                       gc_dbcon.Execute ls_sql
                                     End If
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
                    
                      ls_sql = "INSERT into PO_PayableMaster( Compcode,branchcode, TransCode, TransDate, VendorCode,PaymentType,Bankcode,InstrumentNo,Remarks,userid,adddate,addtime)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtVendorCode & "'," & txtpaymenttype.ListIndex & ",'" & txtbankcode & "','" & txtinstrumentNo & "','" & RepApp(TxtRemarks) & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "' )"
                      gc_dbcon.Execute ls_sql
                
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      If .TextMatrix(ln_cnt, 1) <> "" And Val(.TextMatrix(ln_cnt, 8)) <> 0 Then
                        ls_sql = "INSERT into PO_PayableDetail(Compcode,BranchCode, TransCode, GRNCode, Amount,GSTAmount,SedAmount,DiscAmount,PaidAmount,Remarks)"
                        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "'," & Val(.TextMatrix(ln_cnt, 2)) & "," & Val(.TextMatrix(ln_cnt, 3)) & "," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & "," & Val(.TextMatrix(ln_cnt, 8)) & ",'" & Trim(.TextMatrix(ln_cnt, 9)) & "')"
                        gc_dbcon.Execute ls_sql
                        
                        ls_sql = "Update  PO_POGRN set RecAmount =  " & CheckBalAmount(Trim(.TextMatrix(ln_cnt, 1))) & " WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(.TextMatrix(ln_cnt, 1)) & "'"
                        gc_dbcon.Execute ls_sql
                     End If
                    Next
                 End With
                 
                 
                 
     End Select
gc_dbcon.CommitTrans
'PR_ICIssue.Requery
If Mode <> "D" Then
   'ls_opt = MsgBox("Print Demand Note ?.", vbYesNo)
   'If ls_opt = vbYes Then Call PrintDemandnote
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

Private Function CheckBalAmount(ls_GrnCODE As String) As Double
Dim ln_PaidAmount As Double


pr_dumy1.Open "Select Sum(paidAmount) as Amount from  PO_PayableDetail where compcode = '" & Gs_compcode & "' and GRNCode = '" & ls_GrnCODE & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1

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
     txtVendorCode = Trim(PR_ICIssue("VendorCode") & "")
     If txtVendorCode <> "" Then Call txtVendorCode_KeyDown(vbKeyReturn, vbKeyShift)
     TxtRemarks = Trim(PR_ICIssue("Remarks") & "")
     txtpaymenttype.ListIndex = Val(PR_ICIssue("Paymenttype"))
     txtbankcode = Trim(PR_ICIssue("Bankcode") & "")
     If txtbankcode <> "" Then Call txtbankcode_KeyDown(vbKeyReturn, vbKeyShift)
     txtinstrumentNo = Trim(PR_ICIssue("InstrumentNo") & "")
     txtvchrType = Trim(PR_ICIssue("VchrType") & "")
     txtvchrno = Trim(PR_ICIssue("VoucherNo") & "")
End Sub

Public Function ChkInputs() As Boolean
    If Len(txttransno.Text) = txttransno.MaxLength And Len(txtVendorCode) = txtVendorCode.MaxLength Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function
Private Sub txtVendorCode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtVendorCode <> "" And KeyCode = vbKeyReturn Then
    txtVendorCode = DoPad(txtVendorCode, txtVendorCode.MaxLength)
    
        ls_sql = "Select SupplierCode,Description from Ic_Supplier where compcode = '" & Gs_compcode & "' and SupplierCode = '" & txtVendorCode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Vendor code not found", vbCritical)
                'Cancel = True
            Else
                txtVendordesc = pr_dumy("description")
                If Mode = "A" Then
                    LoadGRNTrans
                    TxtRemarks = "Amount paid to " + Trim(txtVendordesc) + " Payment # =" + txttransno
                End If
                GrdGRN.SetFocus
            End If
         pr_dumy.Close
ElseIf txtVendorCode = "" And KeyCode = vbKeyReturn Then
        Command5_Click
End If
End Sub
Private Sub LoadPaidTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String
ls_sql = " SELECT PO_PayableDetail.GRNCode,PO_PayableDetail.paidAmount, PO_PayableDetail.Amount, PO_PayableDetail.GSTAmount, PO_PayableDetail.SEDAmount,"
ls_sql = ls_sql & " PO_PayableDetail.DiscAmount , PO_PayableDetail.Remarks FROM         PO_PayableMaster INNER JOIN"
ls_sql = ls_sql & " PO_PayableDetail ON PO_PayableMaster.Compcode = PO_PayableDetail.Compcode AND"
ls_sql = ls_sql & " PO_PayableMaster.BranchCode = PO_PayableDetail.BranchCode And PO_PayableMaster.TransCode = PO_PayableDetail.TransCode"
ls_sql = ls_sql & "  where PO_PayableMaster.Compcode = '" & Gs_compcode & "' and PO_PayableMaster.Transcode = '" & txttransno & "'"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("GRNCode") & "")
                .TextMatrix(.Row, 2) = Val(Pr_LoadTrans("Amount"))
                .TextMatrix(.Row, 3) = Val(Pr_LoadTrans("GSTAmount"))
                .TextMatrix(.Row, 4) = Val(Pr_LoadTrans("SEDAmount"))
                .TextMatrix(.Row, 5) = Val(Pr_LoadTrans("DiscAmount"))
                .TextMatrix(.Row, 6) = (Val(.TextMatrix(.Row, 2)) + Val(.TextMatrix(.Row, 3)) + Val(.TextMatrix(.Row, 4))) - Val(.TextMatrix(.Row, 5))
                .TextMatrix(.Row, 7) = Val(Pr_LoadTrans("PaidAmount"))
                .TextMatrix(.Row, 8) = Val(Pr_LoadTrans("PaidAmount"))
                .TextMatrix(.Row, 9) = Trim(Pr_LoadTrans("Remarks") & "")
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

Private Sub txtVendorCode_Validate(Cancel As Boolean)
If txtVendorCode <> "" Then
    txtVendorCode = DoPad(txtVendorCode, txtVendorCode.MaxLength)
    
        ls_sql = "Select SupplierCode,Description from Ic_Supplier where compcode = '" & Gs_compcode & "' and SupplierCode = '" & txtVendorCode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Supplier code not found !!!", vbCritical)
                'Cancel = True
            Else
                txtVendordesc = pr_dumy("description")
                TxtRemarks = "Amount paid to " + Trim(txtVendordesc) + " Invoice# =" + txtInvoice
                txtamount.SetFocus
            End If
         pr_dumy.Close
    
End If
End Sub

Private Sub txtpaymenttype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If txtpaymenttype.ListIndex = 0 Then
    TxtRemarks.SetFocus
Else
    txtbankcode.SetFocus
End If
End If

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
   ElseIf KeyCode = vbKeyReturn And Len(txttransno.Text) = 0 Then
           Command1_Click
   End If
  End Sub
Private Sub LoadGRNTrans()
Dim PR_GRNTrans As New Recordset
Dim Pr_Tax As New Recordset

InitializeGrid
Dim ls_sql As String

ls_sql = " SELECT PO_POGRN.TransCode, PO_POGRN.AccountCode, SUM(PO_POGRNDetail.Amount) AS Amount, SUM(PO_POGRNDetail.GSTAmount) AS Gstamount, "
ls_sql = ls_sql & " SUM(PO_POGRNDetail.SEDAmount) AS SedAmount, SUM(PO_POGRNDetail.DiscAmount) AS DiscAmount"
ls_sql = ls_sql & " FROM PO_POGRN INNER JOIN  PO_POGRNDetail ON PO_POGRN.Compcode = PO_POGRNDetail.Compcode AND PO_POGRN.BranchCode = PO_POGRNDetail.BranchCode AND"
ls_sql = ls_sql & " PO_POGRN.TransCode = PO_POGRNDetail.TransCode "
ls_sql = ls_sql & " where PO_POGRN.Compcode = '" & Gs_compcode & "' and PO_POGRN.AccountCode = '" & txtVendorCode & "' and NetAmount <> RecAmount"
ls_sql = ls_sql & " GROUP BY PO_POGRN.TransCode, PO_POGRN.AccountCode "
PR_GRNTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_GRNTrans.EOF Then
        With GrdGRN
            Do While Not PR_GRNTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_GRNTrans("TransCode") & "")
                .TextMatrix(.Row, 2) = Val(PR_GRNTrans("Amount"))
                .TextMatrix(.Row, 3) = Val(PR_GRNTrans("GSTAmount"))
                .TextMatrix(.Row, 4) = Val(PR_GRNTrans("SEDAmount"))
                .TextMatrix(.Row, 5) = Val(PR_GRNTrans("DiscAmount"))
                .TextMatrix(.Row, 6) = (Val(.TextMatrix(.Row, 2)) + Val(.TextMatrix(.Row, 3)) + Val(.TextMatrix(.Row, 4))) - Val(.TextMatrix(.Row, 5))
                .TextMatrix(.Row, 7) = .TextMatrix(.Row, 6) - CheckBalAmount(.TextMatrix(.Row, 1))
                '.TextMatrix(.Row, 8) = .TextMatrix(.Row, 7)
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
 If KeyCode = vbKeyReturn Then txtVendorCode.SetFocus
End Sub
