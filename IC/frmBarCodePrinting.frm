VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBarCodePrinting 
   Caption         =   "Bar Code Printing"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBarCodePrinting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5880
      MaskColor       =   &H00000000&
      TabIndex        =   13
      Top             =   5085
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   6990
      TabIndex        =   12
      Top             =   5085
      Width           =   1035
   End
   Begin VB.CheckBox ChkExistRec 
      Caption         =   "&Generate New Record and Delete Existing"
      Height          =   315
      Left            =   15
      TabIndex        =   11
      Top             =   5085
      Value           =   1  'Checked
      Width           =   3750
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5460
      Width           =   8115
      _ExtentX        =   14314
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
      Height          =   5070
      Left            =   30
      TabIndex        =   1
      Top             =   -45
      Width           =   8010
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
         Height          =   975
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7965
         Begin VB.CommandButton Command3 
            Caption         =   "< &Previous"
            Height          =   330
            Left            =   6180
            TabIndex        =   16
            Top             =   585
            Width           =   930
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Next >"
            Height          =   330
            Left            =   7125
            TabIndex        =   15
            Top             =   585
            Width           =   795
         End
         Begin VB.TextBox txtnoofcopy 
            Height          =   315
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   9
            Text            =   "1"
            Top             =   585
            Width           =   690
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   3045
            Picture         =   "frmBarCodePrinting.frx":030A
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Start From :"
            Height          =   210
            Left            =   675
            TabIndex        =   8
            Top             =   615
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FA Note # :"
            Height          =   210
            Left            =   735
            TabIndex        =   7
            Top             =   240
            Width           =   810
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
         Height          =   3915
         Left            =   75
         TabIndex        =   14
         Top             =   1020
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   6906
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
         TabIndex        =   10
         Top             =   1440
         Width           =   3945
      End
   End
End
Attribute VB_Name = "frmBarCodePrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Dumy As New Recordset
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
        .FormatString = "Sr# |<Custom Code|<Item Name|<Location"
        .ColWidth(1) = 1600
        .ColWidth(2) = 3000
        .ColWidth(3) = 2500
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

ls_sql = " SELECT PO_FANoteDetail.ItemCode+PO_FANoteDetail.SubCode as Code,IC_Item.Description "
ls_sql = ls_sql & " FROM PO_FANoteDetail INNER JOIN IC_Item ON PO_FANoteDetail.Compcode = IC_Item.Compcode AND PO_FANoteDetail.ItemCode = IC_Item.ItemCode "
ls_sql = ls_sql & "  where PO_FANoteDetail.Compcode = '" & Gs_compcode & "' and PO_FANoteDetail.Transcode = '" & txtTransNo & "'"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("Code") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
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
          PR_IcItem.Open " Select * From Ic_Item Where compcode = '" & Gs_compcode & "' and  CustomCode='" & Trim(.TextMatrix(.Row, 1)) & " ' ", gc_dbcon, adOpenStatic, adLockReadOnly
          
          If PR_IcItem.RecordCount <= 0 Then
              Call MsgBox(Gs_RecNFMsg, vbCritical)
             .TextMatrix(.Row, 1) = ""
             
          Else
             .TextMatrix(.Row, 5) = Trim(PR_IcItem("Itemcode") & "")
             .TextMatrix(.Row, 2) = Trim(PR_IcItem("Description") & "")
             .TextMatrix(.Row, 4) = Val(PR_IcItem("Salecost"))
          End If
         PR_IcItem.Close
       Else
           Call GrdGRN_KeyDown(112, vbKeyShift)
       End If
       .Col = 3
       .CellBackColor = vbHighlight
       ElseIf .Col = 3 Then
           .CellBackColor = vbWindowBackground
           If .TextMatrix(.Row, 3) = "" Then
             Call MsgBox("Enter Quantity!!!", vbCritical)
             Exit Sub
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
     End If
   End With
Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
   If .Col = 1 Or .Col = 3 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
   End If
End With
End If

If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      
      If .Col = 1 Then
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
'On Error GoTo LocalErr
Dim ls_icode As String
Dim ls_Ccode As String
Dim ls_Dcode As String
Dim ls_Scode As String
Dim ls_ivcode As String
Dim ls_Cvcode As String
Dim ls_Dvcode As String
Dim ls_Svcode As String

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

    ls_sql = "delete from IC_BarCodePrinting1"
    gc_dbcon.Execute ls_sql

ln_rsetcnt = Val(txtnoofcopy)
With GrdGRN
For ln_cnt = 1 To .Rows - 1
If Trim(.TextMatrix(ln_cnt, 1)) <> "" Then
        
        For ln_cnt1 = 1 To 1
        ls_icode = "I" & Trim(str(ln_rsetcnt))
        ls_Ccode = "C" & Trim(str(ln_rsetcnt))
        ls_Dcode = "D" & Trim(str(ln_rsetcnt))
        ls_Scode = "S" & Trim(str(ln_rsetcnt))

        ls_infields = ls_infields & "" & ls_icode & ", " & ls_Ccode & ", " & ls_Dcode & ", " & ls_Scode & ","
        
        ls_ivcode = "'M-Tech'"
        ls_Cvcode = "'" & Trim(.TextMatrix(ln_cnt, 1)) & "'"
        ls_Dvcode = "'" & Trim(.TextMatrix(ln_cnt, 2)) & "'"
        ls_Svcode = "'" & "Location." & Trim(.TextMatrix(ln_cnt, 3)) & "'"

        ls_invalues = ls_invalues & "" & ls_ivcode & ", " & ls_Cvcode & ", " & ls_Dvcode & ", " & ls_Svcode & ","
        ln_rsetcnt = ln_rsetcnt + 1
        If ln_rsetcnt = 61 Then
            ls_sql = "insert into IC_BarCodePrinting1 (" & ls_infields & " Compcode )" & " Values (" & ls_invalues & "'" & Gs_compcode & "' )"
            gc_dbcon.Execute ls_sql
            ls_sql = ""
            ls_invalues = ""
            ls_infields = ""
            ln_rsetcnt = 1
        End If
        Next
  End If
 Next
 If ln_rsetcnt < 60 And ln_rsetcnt > 1 Then
    ls_sql = "insert into IC_BarCodePrinting1 (" & ls_infields & " Compcode )" & " Values (" & ls_invalues & "'" & Gs_compcode & "' )"
    gc_dbcon.Execute ls_sql
 End If
End With
End If
  With crrpt
        .ReportFileName = App.Path & Gs_ICRepoPath & "\BarCodePrinting1.RPT"
        .WindowTitle = "" & Me.Caption & ""
        '.Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        '.Formulas(1) = "ReportName = 'Stock Ledger Balance'"
        '.Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Action = 1
    End With

Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
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

Private Sub Command5_Click()
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txtTransNo
        Set PO_DESC = Text1
        Gs_SQL = "SELECT TransCode ,TransDate FROM PO_FANote"
        Gs_OrderBy = "ORDER BY TransCode"
        Gs_OtherPara = " Where compcode = '" & Gs_compcode & "' "
        Gs_OrderBy = "ORDER BY TransCode desc"
        
        MyLookupOLDB.Caption = "Fixed Asset Notes"
        MyLookupOLDB.Show 1
        
        If txtTransNo <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub GrdGRN_Click()
'GrdGRN.CellBackColor = vbHighlight

End Sub

Private Sub GrdGRN_EnterCell()
GrdGRN.CellBackColor = vbHighlight
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then 'Delete Key Pressed
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
'Call txtItemCode_KeyDown(vbKeyReturn, vbKeyShift)
End If
End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Trim(txtTransNo.Text) <> "" Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
         PR_ICIssue.Open "select * from PO_FANote where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                If PR_ICIssue.EOF Then
                   Call MsgBox("Purchase Invoice Not found!!!", vbCritical)
                   txtTransNo.SetFocus
                Else
                   LoadGRNTrans
                End If
            
 ElseIf KeyCode = vbKeyReturn And Trim(txtTransNo.Text) = "" Then
           Command5_Click
 End If
 End Sub
Private Sub Form_Load()
InitializeGrid

End Sub


