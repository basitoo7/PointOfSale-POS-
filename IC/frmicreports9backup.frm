VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmicreport9backup 
   Caption         =   "Bar Code Printing"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports9backup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkOnlyBarCode 
      Caption         =   "Only BarCode"
      Height          =   375
      Left            =   3000
      TabIndex        =   33
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CheckBox ChkExpDateOnly 
      Caption         =   "Experi Date Only"
      Height          =   375
      Left            =   8190
      TabIndex        =   31
      Top             =   5535
      Width           =   1935
   End
   Begin VB.CheckBox ChkPriceCard 
      Caption         =   "Price Cards"
      Height          =   375
      Left            =   4440
      TabIndex        =   30
      Top             =   5505
      Width           =   1710
   End
   Begin VB.CheckBox ChkBakery 
      Caption         =   "Bakery Sticker"
      Height          =   375
      Left            =   6240
      TabIndex        =   29
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CheckBox ChkRsTag 
      Caption         =   "Only Rs. Stickers"
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Top             =   5115
      Width           =   1815
   End
   Begin VB.CheckBox chkprintsticker 
      Caption         =   "Print 4 Columns Stickers Report"
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   5520
      Width           =   3030
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
      Left            =   8505
      MaskColor       =   &H00000000&
      TabIndex        =   17
      Top             =   5085
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   9600
      TabIndex        =   16
      Top             =   5085
      Width           =   1035
   End
   Begin VB.CheckBox ChkExistRec 
      Caption         =   "&Generate New Record and Delete Existing"
      Height          =   315
      Left            =   135
      TabIndex        =   15
      Top             =   5085
      Value           =   1  'Checked
      Width           =   3750
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5955
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   476
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
      Left            =   30
      TabIndex        =   1
      Top             =   -45
      Width           =   10605
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   5760
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   -15
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
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   1365
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   10605
         Begin VB.CommandButton Command4 
            Caption         =   "&Update"
            Height          =   315
            Left            =   9165
            TabIndex        =   24
            Top             =   960
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   9165
            TabIndex        =   22
            Top             =   210
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102039553
            CurrentDate     =   42322
         End
         Begin VB.CommandButton Command3 
            Caption         =   "< &Previous"
            Height          =   330
            Left            =   6180
            TabIndex        =   20
            Top             =   960
            Width           =   930
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Next >"
            Height          =   330
            Left            =   7125
            TabIndex        =   19
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox txtnoofcopy 
            Height          =   315
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   13
            Text            =   "1"
            Top             =   960
            Width           =   690
         End
         Begin VB.TextBox txtitemcode 
            Height          =   315
            Left            =   1680
            TabIndex        =   10
            Top             =   570
            Width           =   1350
         End
         Begin VB.TextBox txtitemdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3390
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   570
            Width           =   4530
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   3045
            Picture         =   "frmicreports9backup.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   555
            Width           =   315
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   3045
            Picture         =   "frmicreports9backup.frx":047C
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3375
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   210
            Width           =   4545
         End
         Begin VB.TextBox txtTransNo 
            Height          =   315
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   4
            Top             =   210
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   9165
            TabIndex        =   23
            Top             =   570
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            Format          =   102039553
            CurrentDate     =   42322
         End
         Begin VB.Label Label6 
            Caption         =   "B.B :"
            Height          =   180
            Left            =   8700
            TabIndex        =   26
            Top             =   615
            Width           =   435
         End
         Begin VB.Label Label5 
            Caption         =   "MFG :"
            Height          =   180
            Left            =   8640
            TabIndex        =   25
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Start From :"
            Height          =   210
            Left            =   675
            TabIndex        =   12
            Top             =   990
            Width           =   840
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Item Code :"
            Height          =   210
            Left            =   720
            TabIndex        =   11
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Invoice No :"
            Height          =   210
            Left            =   705
            TabIndex        =   7
            Top             =   240
            Width           =   840
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   3540
         Left            =   75
         TabIndex        =   18
         Top             =   1425
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   6244
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
      Begin VB.Label Label4 
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
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Top             =   1440
         Width           =   3945
      End
   End
   Begin VB.CheckBox ChkExpDate 
      Caption         =   "Experi Date Sticker"
      Height          =   375
      Left            =   6225
      TabIndex        =   28
      Top             =   5535
      Width           =   1935
   End
End
Attribute VB_Name = "frmicreport9backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Public codeid As String
Public Reporttype As String
Dim ls_sql As String
Dim PR_ICIssue As New Recordset
Dim PR_IcItem As New Recordset
Dim ln_cnt As Integer
Dim ln_cnt1 As Integer
Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String
Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Custom Code|<Item Name|<Qty|<Rate|<Itemcode|<MFG|<B.B|<DiscAmt"
        .ColWidth(1) = 1600
        .ColWidth(2) = 3000
        .ColWidth(3) = 1100
        .ColWidth(4) = 1300
        .ColWidth(5) = 0
        .ColWidth(6) = 1300
        .ColWidth(7) = 1300
        .ColWidth(9) = 500
        
     
        .Redraw = True
         PI_CurRow = 0
         PI_SrNo = 0
    End With
End Sub
Private Sub LoadGRNTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String

ls_sql = " SELECT PO_POGRNDetail.CustomCode,PO_POGRNDetail.ItemCode, IC_Item.salecost as SaleCost,IC_Item.discamt as DiscAmt,IC_Item.Description, PO_POGRNDetail.Quantity+PO_POGRNDetail.BonusQty as Quantity "
ls_sql = ls_sql & " FROM PO_POGRNDetail INNER JOIN IC_Item ON PO_POGRNDetail.Compcode = IC_Item.Compcode AND PO_POGRNDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where PO_POGRNDetail.Compcode = '" & Gs_compcode & "' and PO_POGRNDetail.Transcode = '" & txtTransNo & "'"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("CustomCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Val(Pr_LoadTrans("Quantity"))
                
                .TextMatrix(.Row, 8) = Val(Pr_LoadTrans("DiscAmt"))
                
                .TextMatrix(.Row, 4) = Val(Pr_LoadTrans("SaleCost"))
                
                .TextMatrix(.Row, 5) = Trim(Pr_LoadTrans("Itemcode") & "")
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
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer)
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler

If KeyAscii = 13 Then 'if Enter Key then...
  With argFlexGrid
    If .Col = 1 Then
        .CellBackColor = vbWindowBackground
       If .TextMatrix(.Row, 1) <> "" Then
          If PR_IcItem.State = 1 Then PR_IcItem.Close
          PR_IcItem.Open " Select * From Ic_Item Where compcode = '" & Gs_compcode & "' and  CustomCode='" & Trim(.TextMatrix(.Row, 1)) & "'", gc_dbcon, adOpenStatic, adLockReadOnly
          
          If PR_IcItem.RecordCount <= 0 Then
              Call MsgBox(Gs_RecNFMsg, vbCritical)
             .TextMatrix(.Row, 1) = ""
             
          Else
             .TextMatrix(.Row, 5) = Trim(PR_IcItem("Itemcode") & "")
             .TextMatrix(.Row, 2) = Trim(PR_IcItem("Description") & "")
             .TextMatrix(.Row, 4) = Val(PR_IcItem("Salecost"))
             .TextMatrix(.Row, 8) = Val(PR_IcItem("DiscAmt"))
             .TextMatrix(.Row, 4) = Val(PR_IcItem("Salecost")) - Val(PR_IcItem("DiscAmt"))
          End If
         PR_IcItem.Close
       Else
           Call GrdGRN_KeyDown(112, vbKeyShift)
       End If
  
       .Col = 3
       .CellBackColor = vbHighlight
       'ElseIf .Col = 3 Then
           .CellBackColor = vbWindowBackground
            .TextMatrix(.Row, 3) = 1
           If .TextMatrix(.Row, 3) = "" Then
             Call MsgBox("Enter Quantity!!!", vbCritical)
             Exit Sub
           End If
   End If
         If .TextMatrix(.Row, 1) <> "" Then
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .Row = .Row + 1
           If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
           End If
            
          .SetFocus
         Else
           Call MsgBox("Enter/Select Item Code!!!", vbCritical)
           .Row = .Row
           .Col = 1
         End If
   ' End If
   End With
Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
   If .Col = 1 Or .Col = 3 Or .Col = 6 Or .Col = 7 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
   End If
End With
End If

If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      
      If .Col = 1 Or .Col = 6 Or .Col = 7 Then
        If .CellBackColor = vbHighlight Then
         .Text = "": .CellBackColor = vbWindowBackground
        End If
        .Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
      ElseIf .Col = 3 Then
        If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
        End If
         .Text = .Text & Chr(KeyAscii)
          If Not IsNumeric(.Text) Then
          .Text = ""
           Call MsgBox("Enter Numeric entry !!!", vbCritical)
           Exit Sub
          End If
      End If
    End With
  End If
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
If chkprintsticker.Value = 1 Then
Fourstickersreport
Exit Sub
End If


Dim ls_icode As String
Dim ls_Ccode As String
Dim ls_Dcode As String
Dim ls_Scode As String
Dim ls_ivcode As String
Dim ls_Cvcode As String
Dim ls_Dvcode As String
Dim ls_Svcode As String

Dim ls_mfg As String
Dim ls_edate As String

Dim ls_infields As String
Dim ls_invalues As String
Dim ln_rsetcnt As Integer


If Val(txtnoofcopy) = 0 Then
  Call MsgBox("Enter No of Copies !!!", vbCritical)
  txtnoofcopy.SetFocus
  Exit Sub
End If



If ChkExistRec.Value = 1 Or ChkExpDate.Value = 1 Or ChkBakery.Value = 1 Then
Label4.Caption = "Processing Data Please Wait..."
Me.Refresh

    ls_sql = "delete from IC_BarCodePrinting2"
    gc_dbcon.Execute ls_sql

ln_rsetcnt = Val(txtnoofcopy)
With GrdGRN
For ln_cnt = 1 To .Rows - 1
If Val(.TextMatrix(ln_cnt, 3)) > 0 Then
        
        For ln_cnt1 = 1 To Val(.TextMatrix(ln_cnt, 3))
        ls_icode = "I" & Trim(str(ln_rsetcnt))
        ls_Ccode = "C" & Trim(str(ln_rsetcnt))
        ls_Dcode = "D" & Trim(str(ln_rsetcnt))
        ls_Scode = "S" & Trim(str(ln_rsetcnt))
        
        ls_mfg = "M" & Trim(str(ln_rsetcnt))
        ls_edate = "E" & Trim(str(ln_rsetcnt))
        
        
       ' ls_infields = ls_infields & "" & ls_icode & ", " & ls_Ccode & ", " & ls_Dcode & ", " & ls_Scode & ","
        ls_infields = ls_infields & "" & ls_icode & ", " & ls_Ccode & ", " & ls_Dcode & ", " & ls_Scode & ", " & ls_mfg & ", " & ls_edate & ","
        
        ls_ivcode = "'Rahat Nespak'"
        ls_Cvcode = "'" & Trim(.TextMatrix(ln_cnt, 1)) & "'"
        ls_Dvcode = "'" & RepApp(Trim(.TextMatrix(ln_cnt, 2))) & "'"
        ls_Svcode = "'" & "Rs." & Trim(.TextMatrix(ln_cnt, 4)) & "'"
      
        ls_mfg = "'" & Format(.TextMatrix(ln_cnt, 6), "YYYY/MM/DD") & "'"
        ls_edate = "'" & Format(.TextMatrix(ln_cnt, 7), "YYYY/MM/DD") & "'"
     
        
        'ls_invalues = ls_invalues & "" & ls_ivcode & ", " & ls_Cvcode & "," & ls_Dvcode & "," & ls_Svcode & ","
        ls_invalues = ls_invalues & "" & ls_ivcode & ", " & ls_Cvcode & ", " & ls_Dvcode & ", " & ls_Svcode & ", " & ls_mfg & ", " & ls_edate & ","
        
        
        
        If ln_rsetcnt >= 3 Then
            ls_sql = "insert into IC_BarCodePrinting2 (" & ls_infields & " Compcode )" & " Values (" & ls_invalues & "'" & Gs_compcode & "' )"
            gc_dbcon.Execute ls_sql
            ls_sql = ""
            ls_invalues = ""
            ls_infields = ""
            ln_rsetcnt = 0
       End If
        
        ln_rsetcnt = ln_rsetcnt + 1
       
        Next
  End If
 Next
 
 
 
 If ln_rsetcnt <= 3 Then
    ls_sql = "insert into IC_BarCodePrinting2 (" & ls_infields & " Compcode )" & " Values (" & ls_invalues & "'" & Gs_compcode & "' )"
    gc_dbcon.Execute ls_sql
 End If
End With
'If txtselectedcode <> "" And txtitemcode <> "" Then
'    ls_sql = ls_sql & " where catcode = '" & txtselectedcode & "' and Itemcode = '" & txtitemcode & "'"
'ElseIf txtselectedcode <> "" Then
'    ls_sql = ls_sql & " where catcode = '" & txtselectedcode & "'"
'ElseIf txtitemcode <> "" Then
'    ls_sql = ls_sql & " where Itemcode = '" & txtitemcode & "'"
'End If
'For ln_cnt = 1 To Val(txtnoofcopy)
'gc_dbcon.Execute ls_sql
'Next
'Else
'  Call MsgBox("Enter/Select Category Code !!!", vbCritical)
'  txtselectedcode.SetFocus
'End If
'Label4.Caption = ""
'Me.Refresh
End If
'
With crrpt
          If ChkPriceCard.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrintingPriceCard.rpt"
         ElseIf ChkBakery.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrinting2Col.rpt"
         ElseIf ChkExpDate.Value = 1 Then
           .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrintingExpDate.RPT"
         ElseIf ChkExpDateOnly.Value = 1 Then
           .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrintingExpDateOnly.RPT"
         ElseIf ChkRsTag.Value = 1 Then
           .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrinting2Rs.RPT"
         ElseIf ChkOnlyBarCode.Value = 1 Then
           .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrintingOnly.RPT"
         Else
           .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrinting2.RPT"
         End If
        .WindowTitle = "" & Me.Caption & ""
        '.Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        '.Formulas(1) = "ReportName = 'Stock Ledger Balance'"
        '.Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With

Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Public Function RepApp(str As String) As String
  RepApp = VBA.Replace(str, "'", "''")
End Function

Private Sub Fourstickersreport()
On Error GoTo LocalErr
Dim ls_icode As String
Dim ls_Ccode As String
Dim ls_Dcode As String
Dim ls_Scode As String
Dim ls_ivcode As String
Dim ls_Cvcode As String
Dim ls_Dvcode As String
Dim ls_Svcode As String

Dim ls_mfg As String
Dim ls_edate As String


Dim ls_infields As String
Dim ls_invalues As String
Dim ln_rsetcnt As Integer


If Val(txtnoofcopy) = 0 Then
  Call MsgBox("Enter No of Copies !!!", vbCritical)
  txtnoofcopy.SetFocus
  Exit Sub
End If



If ChkExistRec.Value = 1 Then
Label4.Caption = "Processing Data Please Wait..."
Me.Refresh

    ls_sql = "delete from IC_BarCodePrinting2"
    gc_dbcon.Execute ls_sql

ln_rsetcnt = Val(txtnoofcopy)
With GrdGRN
For ln_cnt = 1 To .Rows - 1
If Val(.TextMatrix(ln_cnt, 3)) > 0 Then
        
        For ln_cnt1 = 1 To Val(.TextMatrix(ln_cnt, 3))
        ls_icode = "I" & Trim(str(ln_rsetcnt))
        ls_Ccode = "C" & Trim(str(ln_rsetcnt))
        ls_Dcode = "D" & Trim(str(ln_rsetcnt))
        ls_Scode = "S" & Trim(str(ln_rsetcnt))
        ls_mfg = "M" & Trim(str(ln_rsetcnt))
        ls_edate = "E" & Trim(str(ln_rsetcnt))

        ls_infields = ls_infields & "" & ls_icode & ", " & ls_Ccode & ", " & ls_Dcode & ", " & ls_Scode & ", " & ls_mfg & ", " & ls_edate & ","
        
        ls_ivcode = "'Rahat Bakers & Departmental Store'"
        ls_Cvcode = "'" & Trim(.TextMatrix(ln_cnt, 1)) & "'"
        ls_Dvcode = "'" & Trim(.TextMatrix(ln_cnt, 2)) & "'"
        ls_Svcode = "'" & "Rs." & Trim(.TextMatrix(ln_cnt, 4)) & "'"
        ls_mfg = "'" & Format(.TextMatrix(ln_cnt, 6), "YYYY/MM/DD") & "'"
        ls_edate = "'" & Format(.TextMatrix(ln_cnt, 7), "YYYY/MM/DD") & "'"
        

        ls_invalues = ls_invalues & "" & ls_ivcode & ", " & ls_Cvcode & ", " & ls_Dvcode & ", " & ls_Svcode & ", " & ls_mfg & ", " & ls_edate & ","
       
        
        
        
        If ln_rsetcnt >= 4 Then
            ls_sql = "insert into IC_BarCodePrinting2 (" & ls_infields & " Compcode )" & " Values (" & ls_invalues & "'" & Gs_compcode & "' )"
            gc_dbcon.Execute ls_sql
            ls_sql = ""
            ls_invalues = ""
            ls_infields = ""
            ln_rsetcnt = 0
       End If
        
        ln_rsetcnt = ln_rsetcnt + 1
       
        Next
  End If
 Next
 
 
 
 If ln_rsetcnt <= 4 Then
    ls_sql = "insert into IC_BarCodePrinting2 (" & ls_infields & " Compcode )" & " Values (" & ls_invalues & "'" & Gs_compcode & "' )"
    gc_dbcon.Execute ls_sql
 End If
End With
'If txtselectedcode <> "" And txtitemcode <> "" Then
'    ls_sql = ls_sql & " where catcode = '" & txtselectedcode & "' and Itemcode = '" & txtitemcode & "'"
'ElseIf txtselectedcode <> "" Then
'    ls_sql = ls_sql & " where catcode = '" & txtselectedcode & "'"
'ElseIf txtitemcode <> "" Then
'    ls_sql = ls_sql & " where Itemcode = '" & txtitemcode & "'"
'End If
'For ln_cnt = 1 To Val(txtnoofcopy)
'gc_dbcon.Execute ls_sql
'Next
'Else
'  Call MsgBox("Enter/Select Category Code !!!", vbCritical)
'  txtselectedcode.SetFocus
'End If
'Label4.Caption = ""
'Me.Refresh
End If
'
With crrpt
       
        .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrinting4.RPT"
        .WindowTitle = "" & Me.Caption & ""
        '.Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        '.Formulas(1) = "ReportName = 'Stock Ledger Balance'"
        '.Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With

Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0

End Sub
Private Sub Command1_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtitemcode
    Set PO_DESC = txtitemdesc
    Gs_SQL = "SELECT Customcode as Alias, Description,SaleCost,DiscAmt,itemcode  from IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by itemcode"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    
    MyLookupOLDB.Caption = "Items "
    MyLookupOLDB.Show 1
   If txtitemcode <> "" Then Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command2_Click()
 txtTransNo.Text = Val(txtTransNo.Text) + 1
 txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
 If txtTransNo <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command3_Click()
 txtTransNo.Text = Val(txtTransNo.Text) - 1
 txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
 If txtTransNo <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command4_Click()
With GrdGRN
For ln_cnt = 1 To .Rows - 1
.TextMatrix(ln_cnt, 6) = DTPicker1
.TextMatrix(ln_cnt, 7) = DTPicker2

Next
End With
End Sub

Private Sub Command5_Click()
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txtTransNo
        Set PO_DESC = Text1
        Gs_SQL = "SELECT GRN.TransCode AS ComputerCode, GRN.GRNCode AS GRNCode, Vendors.Description AS 'Vendors.Description', GRN.TransDate AS GRNDate,    GRN.NetAmount AS 'GRN.NetAmount' FROM         PO_POGRN GRN INNER JOIN         IC_Supplier Vendors ON GRN.Compcode = Vendors.Compcode AND GRN.AccountCode = Vendors.SupplierCode"
        Gs_OrderBy = "ORDER BY GRN.TransCode"
        Gs_OtherPara = " Where GRN.compcode = '" & Gs_compcode & "' "
        Gs_OrderBy = "ORDER BY GRN.TransCode desc"
        
        frmPosearchRecords.Caption = "GRN"
        frmPosearchRecords.Show 1
        
        If txtTransNo <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub GrdGRN_Click()

GrdGRN.CellBackColor = vbHighlight

End Sub

Private Sub GrdGRN_EnterCell()
GrdGRN.CellBackColor = vbHighlight
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 And Trim(GrdGRN.TextMatrix(GrdGRN.Row, 1)) = "" Then  ' Enter key pressed
    Text1 = ""
    Text2 = ""
If Trim(GrdGRN.TextMatrix(GrdGRN.Row, 1)) = "" Then
    DoEvents
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    Gs_SQL = "SELECT customCode,Description,SaleCost as SalePrice,DiscAmt as DiscAmt,(SaleCost-DiscAmt) as AfterDisc FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    Gs_ExtraPara = " And Customcode = '1'"
    MyLookupSaleOLDB.Caption = "Items"
    MyLookupSaleOLDB.Show 1
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = Text1
    GrdGRN.TextMatrix(GrdGRN.Row, 2) = Text2
End If
   ' Set PO_AnyForm = Nothing
   ' Set PO_AnyForm = Me
   ' Set PO_CODE = Text2
   ' Set PO_DESC = Text3
  '  Gs_SQL = "SELECT customCode,Description,SaleCost,DiscAmt,Itemcode FROM IC_Item "
  '  Gs_FindFld = "Description"
  '  Gs_OrderBy = "Order by Description"
  '  Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
  '  MyLookupOLDB.Caption = "Items Information"
  '  MyLookupOLDB.Show 1
      
    GrdGRN.CellBackColor = vbWindowBackground
     GrdGRN.Col = 1
     GrdGRN.CellBackColor = vbHighlight
      If GrdGRN.TextMatrix(GrdGRN.Row, 1) <> "" Then
        Call GrdGRN_KeyPress(13)
     End If
     
ElseIf KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
    End With
End If
End Sub

Private Sub GrdGRN_KeyPress(KeyAscii As Integer)
  Call GetKeysAdd(GrdGRN, KeyAscii)
  Exit Sub
End Sub

Private Sub GrdGRN_LeaveCell()
With GrdGRN
 .CellBackColor = vbWindowBackground
End With
End Sub

Private Sub txtnoofcopy_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)
End If
End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Trim(txtTransNo.Text) <> "" Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
         PR_ICIssue.Open "select * from PO_POGRN where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                If PR_ICIssue.EOF Then
                   Call MsgBox("Purchase Invoice Not found!!!", vbCritical)
                   txtTransNo.SetFocus
                Else
                   LoadGRNTrans
                End If
            
 ElseIf KeyCode = vbKeyReturn And Trim(txtTransNo.Text) = "" Then
           Command1_Click
 End If
 End Sub




Private Sub Form_Load()
InitializeGrid
DTPicker1 = Date
DTPicker2 = Date + 30
End Sub

Private Sub txtItemcode_Change()
If txtitemcode = "" Then
txtitemdesc = ""
End If

End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtitemcode <> "" And KeyCode = vbKeyReturn Then
    txtitemcode = DoPad(txtitemcode, txtitemcode.MaxLength)
    ls_sql = "Select customcode,salecost-discamt as SaleCost,itemcode,Description from IC_Item where compcode = '" & Gs_compcode & "' and customcode = '" & txtitemcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Item Code not found", vbCritical)
            Else
                txtitemdesc = pr_dumy("description")
                AddToGrid
                txtnoofcopy.SetFocus
            End If
         pr_dumy.Close

End If
End Sub
Private Sub AddToGrid()
Dim ln_cnt As Integer
                    If PS_RowClicked = "" Then
                        If PI_SrNo = 0 Then
                            PI_SrNo = 1
                        Else
                            PI_SrNo = PI_SrNo + 1
                         End If
                     End If
        
                        With GrdGRN
                            If PS_RowClicked = "" Then
                                    If Not PI_SrNo = 1 Then .Rows = .Rows + 1
                                        .Row = .Rows - 1
                                    Else
                                        .Row = PI_CurRow
                                    End If
                                    If PS_RowClicked = "" Then
                                        .TextMatrix(.Row, 0) = PI_SrNo
                                    Else
                                        .TextMatrix(.Row, 0) = PI_CurRow
                                    End If
                                    .TextMatrix(.Row, 1) = pr_dumy("customcode")
                                    .TextMatrix(.Row, 2) = pr_dumy("Description")
                                    .TextMatrix(.Row, 3) = 1
                                    .TextMatrix(.Row, 4) = pr_dumy("Salecost")
                                    .TextMatrix(.Row, 5) = pr_dumy("Itemcode")
                                PS_RowClicked = ""
                        End With
                        
                   
End Sub



