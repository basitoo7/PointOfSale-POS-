VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSO_PosformEdit 
   Caption         =   "POINT OF SALE"
   ClientHeight    =   8190
   ClientLeft      =   -1995
   ClientTop       =   390
   ClientWidth     =   11880
   Icon            =   "frmPosDataFormEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   8595
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   14100
      Begin VB.TextBox txtitemdesc 
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8130
         Width           =   10860
      End
      Begin VB.Frame Frame6 
         Height          =   930
         Left            =   4995
         TabIndex        =   22
         Top             =   0
         Width           =   3555
         Begin VB.CheckBox ChkEmpBill 
            Caption         =   "Employee Bill"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   75
            TabIndex        =   23
            Top             =   300
            Width           =   1950
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   5025
         Begin VB.CommandButton CmdAccount1 
            Height          =   315
            Left            =   3120
            Picture         =   "frmPosDataFormEdit.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   135
            Width           =   315
         End
         Begin MSComCtl2.DTPicker TXTINVDATE 
            Height          =   330
            Left            =   1530
            TabIndex        =   26
            Top             =   510
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   582
            _Version        =   393216
            Format          =   59899905
            CurrentDate     =   40754
         End
         Begin VB.TextBox TXTINVNUMBER 
            Height          =   315
            Left            =   1530
            TabIndex        =   25
            Top             =   150
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "INVOICE DATE :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   300
            TabIndex        =   18
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "INVOICE NO :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            TabIndex        =   17
            Top             =   165
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   930
         Left            =   8520
         TabIndex        =   13
         Top             =   0
         Width           =   2340
         Begin VB.CommandButton Command1 
            Caption         =   "&Hold"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   180
            TabIndex        =   15
            Top             =   255
            Width           =   990
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Restore"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1230
            TabIndex        =   14
            Top             =   255
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   945
         Left            =   10830
         TabIndex        =   10
         Top             =   -15
         Width           =   3270
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Casher :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   15
            TabIndex        =   12
            Top             =   315
            Width           =   765
         End
         Begin VB.Label lblcasherName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   11
            Top             =   315
            Width           =   2280
         End
      End
      Begin VB.TextBox txtempDisc 
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
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4635
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8130
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3915
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   255
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtdiscamount 
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
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2925
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8130
         Visible         =   0   'False
         Width           =   1635
      End
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   12300
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8130
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3510
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   225
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Frame Frame5 
         Height          =   7110
         Left            =   30
         TabIndex        =   19
         Top             =   900
         Width           =   14025
         Begin VB.TextBox TXTBARCODE 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   660
            Locked          =   -1  'True
            MaxLength       =   13
            TabIndex        =   20
            Top             =   1335
            Width           =   2310
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexPOS 
            Height          =   6825
            Left            =   60
            TabIndex        =   21
            Top             =   150
            Width           =   13830
            _ExtentX        =   24395
            _ExtentY        =   12039
            _Version        =   393216
            RowHeightMin    =   400
            BackColorSel    =   16777215
            ForeColorSel    =   0
            GridColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Label Label11 
         Caption         =   " Total Amount :"
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
         Height          =   255
         Left            =   10920
         TabIndex        =   5
         Top             =   8160
         Width           =   1455
      End
      Begin VB.Label txtstatus 
         Height          =   165
         Left            =   4230
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "GRAND TOTAL"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   1
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Label TXTINVTOTAL 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   0
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "FILE"
      Begin VB.Menu MNUEXIT 
         Caption         =   "EXIT"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu MNUDATA 
      Caption         =   "DATA"
      Begin VB.Menu MNUNEWINVOICE 
         Caption         =   "NEW BILL"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "frmSO_PosformEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public j, i
Public X, X2, lin
Public startdate, enddate As Date
Public RS As New Recordset
Public myRS As New Recordset
Public Answer
Public nF, Inp1, Z, b, P, c
Public Totrec
Public CY As Integer
Public CZ As Integer
Public Cz4, CX As Integer
Dim PR_Dumy As New Recordset
Dim ls_sql As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_ICIssue As New Recordset


Private Sub InitializeGrid()
    With MSFlexPOS
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Custom Code|<Description|<QTY|<Sale Price|<Amount|<Category|<U.O.M|<Discperc|<DiscAmount|<Itemcode|<EmpDiscAmount"
        .ColWidth(1) = 2500
        .ColWidth(2) = 3300
        .ColWidth(3) = 900
        .ColAlignment(3) = 7

        .ColWidth(4) = 1200
        .ColAlignment(4) = 7
        
        .ColWidth(5) = 1500
        .ColAlignment(5) = 7

        .ColWidth(6) = 2000
        .ColWidth(7) = 1500
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .Redraw = True
        TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
        TXTBARCODE.Text = .TextMatrix(.Row, 1)
        TXTBARCODE.Visible = True

    End With

End Sub

Private Sub Command1_Click()
If MSFlexPOS.Rows > 2 Then
Dim ls_transcodehold As String
Dim ln_cnt As Integer
ls_transcodehold = maxtranscodehold
 With MSFlexPOS
       For ln_cnt = 1 To .Rows - 1
       If .TextMatrix(ln_cnt, 1) <> "" Then
        ls_sql = "INSERT into SO_TransHold(Compcode, TransCode,transdate,customcode, ItemCode, Quantity,Itemrate,Amount)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & ls_transcodehold & "', '" & Format(TXTINVDATE, "YYYY/MM/DD HH:MM:SS") & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 10)) & "'," & (Val(0 & .TextMatrix(ln_cnt, 3))) & "," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & ")"
        gc_dbcon.Execute ls_sql
      End If
      Next
  End With

TXTINVDATE = Now
TXTINVNUMBER = maxtranscode
InitializeGrid
MSFlexPOS.Row = 1
TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
TXTBARCODE.Visible = True
TXTBARCODE.TabIndex = 0
txttotalamount = ""
TXTBARCODE.Locked = False
TXTBARCODE.SetFocus
Else
    Call MsgBox("Nothing for Hold!!!", vbCritical)
End If
End Sub

Private Sub Command2_Click()
  ls_sql = "Select * from so_Transhold where compcode = '" & Gs_compcode & "'"
  PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
  If PR_Dumy.EOF Then
  Call MsgBox("Nothing for restore !!!", vbCritical)
  PR_Dumy.Close
  TXTBARCODE.SetFocus
  Exit Sub
  End If
  PR_Dumy.Close
  Set PO_AnyForm = Nothing
  Set PO_AnyForm = Me
  Set PO_CODE = Text1
  Set PO_DESC = Text2
  Gs_SQL = "Select TransCode, TransDate from So_Transhold "
  Gs_FindFld = "Transcode"
  Gs_OrderBy = "Order by Transdate"
  Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' group by Transcode,Transdate "
  MyLookupOLDB.Caption = "Hold Trans"
  MyLookupOLDB.Show 1
  InitializeGrid
  If Text1 <> "" Then

    ls_sql = "SELECT IC_Item.CustomCode, IC_Item.ItemCode, IC_Item.Description, SO_TransHold.Quantity, SO_TransHold.ItemRate, SO_TransHold.Amount,"
    ls_sql = ls_sql & " IC_Item.SaleDiscPerc, IC_Item.SaleCost, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc,IC_ItemCategory.empdiscper FROM IC_Item INNER JOIN"
    ls_sql = ls_sql & " IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN"
    ls_sql = ls_sql & " IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode INNER JOIN"
    ls_sql = ls_sql & " SO_TransHold ON IC_Item.Compcode = SO_TransHold.Compcode AND IC_Item.ItemCode = SO_TransHold.ItemCode"
    ls_sql = ls_sql & " WHERE (SO_TransHold.Compcode = '" & Gs_compcode & "') AND (SO_TransHold.TransCode = '" & Text1 & "')"
End If
With MSFlexPOS
PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
            Do While Not PR_Dumy.EOF
               .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 1) = Trim(PR_Dumy("Customcode") & "")
               .TextMatrix(.Row, 2) = Trim(PR_Dumy("Description") & "")
               .TextMatrix(.Row, 3) = Val(0 & PR_Dumy("Quantity"))
               .TextMatrix(.Row, 4) = Val(0 & PR_Dumy("ItemRate"))
               .TextMatrix(.Row, 5) = Val(0 & PR_Dumy("Amount"))
               .TextMatrix(.Row, 6) = Trim(PR_Dumy("CatDesc") & "")
               .TextMatrix(.Row, 7) = Trim(PR_Dumy("UOM") & "")
               .TextMatrix(.Row, 8) = Val(0 & PR_Dumy("SaleDiscPerc"))
               .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 5)) * Val(0 & PR_Dumy("SaleDiscPerc")) / 100
               .TextMatrix(.Row, 10) = Trim(PR_Dumy("Itemcode") & "")
               .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(0 & PR_Dumy("EmpDiscPer")) / 100
        If .Row = .Rows - 1 Then
        .Col = 1
        .Row = .Rows - 1
        .Rows = .Rows + 1
        .Row = .Rows - 1
         TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
         TXTBARCODE.Text = .TextMatrix(.Row, 1)
         TXTBARCODE.Visible = True
        ElseIf .Row < .Rows - 1 Then
            .Row = .Rows - 1
             TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
             TXTBARCODE.Text = .TextMatrix(.Row, 1)
        End If
        TotalAmount
        
      PR_Dumy.MoveNext
      Loop
Else
    MsgBox "Hold Transaction not found", vbExclamation, "Error"
    InitializeGrid
    TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
End If
End With
PR_Dumy.Close
ls_sql = "Delete from So_TransHold where Transcode = '" & Text1 & "' and compcode = '" & Gs_compcode & "'"
gc_dbcon.Execute ls_sql
TXTINVDATE = Now
TXTINVNUMBER = maxtranscode
TXTBARCODE.SetFocus
End Sub

Private Sub Form_Click()
'InitializeGrid
'TXTINVDATE = Date
'TXTBARCODE.Visible = False
End Sub

Private Sub mnuNewinvoice_Click()
If Val(txttotalamount) > 0 Then
    txtstatus = ""
    frmSOPaidAmtform.txttotalamount = txttotalamount
    
    If ChkEmpBill.Value = 1 Then
        frmSOPaidAmtform.txtdisamount = Val(txtempDisc)
    End If
    
    frmSOPaidAmtform.Show 1
End If
If txtstatus <> "Cancel" Then
TXTINVDATE = Now
TXTINVNUMBER = maxtranscode
InitializeGrid
MSFlexPOS.Row = 1
TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
TXTBARCODE.Visible = True
TXTBARCODE.TabIndex = 0
txttotalamount = ""
'RS.Open "select max(invno) + 1 from pos_tbl", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'TXTINVNUMBER.Caption = Format(RS.Fields(0), "0")
'RS.Close

TXTBARCODE.Locked = False
End If
End Sub

Private Sub TotalAmount()
    Dim ln_cnt As Integer
      txttotalamount = ""
      txtdiscamount = ""
      txtempDisc = ""
    With MSFlexPOS
        For ln_cnt = 1 To .Rows - 1
            txttotalamount = Format(Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
            txtdiscamount = Format(Val(txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), "######0.00")
            txtempDisc = Format(Val(txtempDisc) + Val(.TextMatrix(ln_cnt, 11)), "######0.00")
        Next
    End With
End Sub
Private Sub MSFLEXPOS_EnterCell()
With MSFlexPOS
.CellBackColor = vbHighlight
End With
End Sub

Private Sub MSFlexPOS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With MSFlexPOS
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
            TotalAmount
    End With
 End If
End Sub

Private Sub MSFlexPOS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MSFlexPOS.Row = MSFlexPOS.Rows - 1
    MSFlexPOS.Col = 1
    If TXTBARCODE.Visible = False Then
    With MSFlexPOS
        TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
        TXTBARCODE.Text = .TextMatrix(.Row, 1)
        TXTBARCODE.Visible = True
        TXTBARCODE.TabIndex = 0
        TXTBARCODE.SetFocus
    End With
    End If
    
    Exit Sub
End If
With MSFlexPOS
    If .Col = 3 Then
            If KeyAscii <> 27 And KeyAscii <> 8 Then
              If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
              End If
             .Text = .Text & Chr(KeyAscii)
            End If
            .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 3)) * Val(.TextMatrix(.Row, 4))
             TotalAmount
    End If
End With

If KeyAscii = 8 Then  'If BackSpace Key then...
With MSFlexPOS
   If .Col = 1 Or .Col = 3 Or .Col = 4 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
       .CellBackColor = vbWindowBackground
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
  End If
End With
End If

End Sub

Private Sub MSFLEXPOS_LeaveCell()
MSFlexPOS.CellBackColor = vbWindowBackground
MSFlexPOS.SelectionMode = flexSelectionFree
End Sub

    Public Sub BGLoss(anyControl As Control)
        If anyControl.Locked = False Then
          anyControl.BackColor = vbYellow
        End If
    End Sub
    Public Sub HFocus(ByRef sText As Variant)
        With sText
            .SelStart = 0
            .SelLength = Len(sText.Text)
        End With
    End Sub
Public Sub DISAPEARME()
    Dim s
    For Each s In Controls
    If TypeOf s Is TextBox Then
    s.Visible = False
    End If
    Next
End Sub

Public Sub BGFocus(anyControl As Control)
        If anyControl.Locked = False Then
            'anyControl.Backcolor = &H80000013
            anyControl.BackColor = vbGreen
        End If
End Sub

Private Sub Form_Load()
InitializeGrid
TXTINVDATE = Now
lblcasherName = Gc_UserName
With MSFlexPOS
If .Col = 1 Then
 TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
 TXTBARCODE.Text = .TextMatrix(.Row, 1)
 TXTBARCODE.Visible = True
 TXTBARCODE.TabIndex = 0
End If
End With
End Sub
Private Sub LOADHEADER()

With MSFlexPOS
'.RowHeight(0) = 500
    .Cols = 12
    .Rows = 2
    .ColWidth(0) = 800:
    .ColWidth(1) = 2500:
    .ColWidth(2) = 3300:
    .ColWidth(3) = 1500:
    .ColWidth(4) = 1100:
    .ColWidth(0) = 600:
    .ColWidth(7) = 1500:
    .ColWidth(10) = 1600:
    

    

    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignLeftCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColAlignment(6) = flexAlignRightCenter
    .ColAlignment(7) = flexAlignRightCenter
    .ColAlignment(8) = flexAlignRightCenter
    .ColAlignment(9) = flexAlignRightCenter
    .ColAlignment(10) = flexAlignRightCenter
    .ColAlignment(11) = flexAlignRightCenter
    

            .TextMatrix(0, 0) = "No"
            .TextMatrix(0, 1) = "Customcode"
            .TextMatrix(0, 2) = "Description"
            .TextMatrix(0, 3) = "Qty"
            .TextMatrix(0, 4) = "Sale Price"
            .TextMatrix(0, 7) = "Sub Total"
            .TextMatrix(0, 3) = "Category"
            .TextMatrix(0, 4) = "U.O.M"
            
            
            .TextMatrix(0, 8) = "Vat %"
            .TextMatrix(0, 9) = "Vat amount"
            .TextMatrix(0, 10) = "Line Total"
            .TextMatrix(0, 11) = "Row No"
        End With
        
        
    End Sub


Private Sub MNUEXIT_Click()
Unload Me
End Sub

Private Sub MSFlexPOS_Click()
With MSFlexPOS
If .Col = 1 Then
'TXTBARCODE.Text = UCase(TXTBARCODE.Text)
TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
TXTBARCODE.Text = .TextMatrix(.Row, 1)
TXTBARCODE.Visible = True
TXTBARCODE.SetFocus
.CellBackColor = vbHighlight
MSFlexPOS.SelectionMode = flexSelectionFree

End If
txtitemdesc = .TextMatrix(.Row, 2)
End With
End Sub

Public Function maxtranscode() As String
PR_Dumy.Open "select max(transcode) as transcode from SO_TransMaster where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & PR_Dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
PR_Dumy.Close
End Function
Public Function maxtranscodehold() As String
PR_Dumy.Open "select max(transcode) as transcode from SO_Transhold where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
    maxtranscodehold = DoPad(Trim(str(Int(0 & PR_Dumy("transcode")) + 1)), 10)
Else
    maxtranscodehold = DoPad(Trim(str(Int(1))), 10)
End If
PR_Dumy.Close
End Function


Private Sub TXTBARCODE_DblClick()
With MSFlexPOS
If .Col = 1 Then
          TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
             TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
               TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
                'TXTPARTICULARS.Visible = False
                TXTBARCODE.Visible = True
                TXTBARCODE.SetFocus
                End If
            .CellBackColor = vbHighlight
    End With
End Sub

Public Sub HighLight()
        With Screen.ActiveForm
            If (TypeOf .ActiveControl Is TextBox) Then
                .ActiveControl.SelStart = 0
                .ActiveControl.SelLength = Len(.ActiveControl)
    '         .ActiveControl.Backcolor = &H80000013
            .ActiveControl.BackColor = vbYellow
            End If
        End With
End Sub

Private Sub TXTBARCODE_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 And MSFlexPOS.Col = 1 Then  ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TXTBARCODE
    Set PO_DESC = Text2
    Gs_SQL = "SELECT customCode,Description"
    Gs_SQL = Gs_SQL & " FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
 ElseIf KeyCode = 13 And MSFlexPOS.Col = 1 And Trim(TXTBARCODE) = "" Then   ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TXTBARCODE
    Set PO_DESC = Text2
    Gs_SQL = "SELECT customCode,Description"
    Gs_SQL = Gs_SQL & " FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
ElseIf KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With MSFlexPOS
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
            TotalAmount
    End With
    TXTBARCODE.Visible = False
 End If


If Trim(TXTBARCODE) <> "" Then
    Call TXTBARCODE_KeyPress(13)
End If

End Sub
Private Sub TXTBARCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(TXTBARCODE) <> "" Then

With MSFlexPOS
ls_sql = "SELECT IC_Item.CustomCode,IC_Item.ItemCode, IC_Item.Description, IC_Item.SaleDiscPerc, IC_Item.SaleCost, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc,IC_ItemCategory.EmpDiscPer"
ls_sql = ls_sql & " FROM IC_Item INNER JOIN IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode"
ls_sql = ls_sql & " where IC_Item.compcode = '" & Gs_compcode & "' and IC_Item.Customcode  = '" & Trim(TXTBARCODE) & "'"

PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1

If Not PR_Dumy.EOF Then

               .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 1) = Trim(PR_Dumy("Customcode") & "")
               .TextMatrix(.Row, 2) = Trim(PR_Dumy("Description") & "")
               txtitemdesc = .TextMatrix(.Row, 2)
               .TextMatrix(.Row, 3) = 1
               .TextMatrix(.Row, 4) = Val(0 & PR_Dumy("Salecost"))
               .TextMatrix(.Row, 5) = Val(0 & PR_Dumy("Salecost")) * .TextMatrix(.Row, 3)
               .TextMatrix(.Row, 6) = Trim(PR_Dumy("CatDesc") & "")
               .TextMatrix(.Row, 7) = Trim(PR_Dumy("UOM") & "")
               .TextMatrix(.Row, 8) = Val(0 & PR_Dumy("SaleDiscPerc"))
               .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 5)) * Val(0 & PR_Dumy("SaleDiscPerc")) / 100
               .TextMatrix(.Row, 10) = Trim(PR_Dumy("Itemcode") & "")
               .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(0 & PR_Dumy("EmpDiscPer")) / 100
      
        If .Row = .Rows - 1 Then
        .Col = 1
        .Row = .Rows - 1
        .Rows = .Rows + 1
        .Row = .Rows - 1
         TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
         TXTBARCODE.Text = .TextMatrix(.Row, 1)
         TXTBARCODE.Visible = True
        ElseIf .Row < .Rows - 1 Then
             Caption = "edit"
            .Row = .Rows - 1
             TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
             TXTBARCODE.Text = .TextMatrix(.Row, 1)
        End If
        TotalAmount
 
Else
    MsgBox "Item Code not found", vbExclamation, "Error"


End If
    PR_Dumy.Close
    TXTBARCODE.Visible = False
    .Col = 1
    TXTBARCODE.Visible = True
If TXTBARCODE.Enabled Then TXTBARCODE.SetFocus
End With
End If
End Sub




Private Sub TXTINVNUMBER_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Len(TXTINVNUMBER.Text) > 0 Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         TXTINVNUMBER.Text = DoPad(UCase(TXTINVNUMBER.Text), 10)
         PR_ICIssue.Open "select * from SO_TransMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & TXTINVNUMBER & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         If PR_ICIssue.EOF Then
                   Call MsgBox("Record not found !!!", vbCritical)
                   TXTINVNUMBER.SetFocus
         Else
                   TXTINVDATE = PR_ICIssue("Transdate")
                   InitializeGrid
                   LoadGRNTrans
         End If
End If
End Sub
Private Sub LoadGRNTrans()
ls_sql = "SELECT IC_Item.CustomCode, IC_Item.ItemCode, IC_Item.Description, SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount,"
ls_sql = ls_sql & " IC_Item.SaleDiscPerc, IC_Item.SaleCost, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc FROM IC_Item INNER JOIN"
ls_sql = ls_sql & " IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN"
ls_sql = ls_sql & " IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode INNER JOIN"
ls_sql = ls_sql & " SO_Trans ON IC_Item.Compcode = SO_Trans.Compcode AND IC_Item.ItemCode = SO_Trans.ItemCode"
ls_sql = ls_sql & " WHERE (SO_Trans.Compcode = '" & Gs_compcode & "') AND (SO_Trans.TransCode = '" & TXTINVNUMBER & "')"

With MSFlexPOS
PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
            Do While Not PR_Dumy.EOF
               .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 1) = Trim(PR_Dumy("Customcode") & "")
               .TextMatrix(.Row, 2) = Trim(PR_Dumy("Description") & "")
               .TextMatrix(.Row, 3) = Val(0 & PR_Dumy("Quantity"))
               .TextMatrix(.Row, 4) = Val(0 & PR_Dumy("ItemRate"))
               .TextMatrix(.Row, 5) = Val(0 & PR_Dumy("Amount"))
               .TextMatrix(.Row, 6) = Trim(PR_Dumy("CatDesc") & "")
               .TextMatrix(.Row, 7) = Trim(PR_Dumy("UOM") & "")
               .TextMatrix(.Row, 8) = Val(0 & PR_Dumy("SaleDiscPerc"))
               .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 5)) * Val(0 & PR_Dumy("SaleDiscPerc")) / 100
                .TextMatrix(.Row, 10) = Trim(PR_Dumy("Itemcode") & "")
        If .Row = .Rows - 1 Then
        .Col = 1
        .Row = .Rows - 1
        .Rows = .Rows + 1
        .Row = .Rows - 1
         TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
         TXTBARCODE.Text = .TextMatrix(.Row, 1)
         TXTBARCODE.Visible = True
        ElseIf .Row < .Rows - 1 Then
            .Row = .Rows - 1
             TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
             TXTBARCODE.Text = .TextMatrix(.Row, 1)
        End If
        TotalAmount
        
      PR_Dumy.MoveNext
      Loop
Else
    MsgBox "Transaction not found", vbExclamation, "Error"
    InitializeGrid
    TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
End If
End With
PR_Dumy.Close

TXTBARCODE.SetFocus
End Sub


