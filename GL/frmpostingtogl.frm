VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frmpostingtogl 
   Caption         =   "Posting To Gl"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   Icon            =   "frmpostingtogl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Tag             =   " v      "
   Begin VB.CommandButton Command2 
      Caption         =   "Post Selected Voucher"
      Height          =   360
      Left            =   8190
      TabIndex        =   14
      Top             =   6330
      Width           =   2025
   End
   Begin VB.Frame Frame2 
      Height          =   6225
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   10155
      Begin VB.CheckBox chkselectall 
         Caption         =   "Select All"
         Height          =   330
         Left            =   4890
         TabIndex        =   16
         Top             =   585
         Width           =   1620
      End
      Begin VB.TextBox txtvoucherno 
         Height          =   300
         Left            =   8160
         TabIndex        =   13
         Top             =   5655
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtvchrtype 
         Height          =   300
         Left            =   7395
         TabIndex        =   12
         Top             =   5655
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtuserid 
         Height          =   330
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   11
         Top             =   195
         Width           =   1860
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load Vouchers"
         Height          =   300
         Left            =   3285
         TabIndex        =   10
         Top             =   585
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker DTPvoucherdate 
         Height          =   315
         Left            =   1425
         TabIndex        =   9
         Top             =   585
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16318465
         CurrentDate     =   39071
      End
      Begin VB.PictureBox picUnchecked 
         Height          =   285
         Left            =   0
         Picture         =   "frmpostingtogl.frx":030A
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picChecked 
         Height          =   285
         Left            =   15
         Picture         =   "frmpostingtogl.frx":064C
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   7
         Top             =   105
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmdLookup 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         Picture         =   "frmpostingtogl.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox txtuseriddesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   3645
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   6405
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   7020
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   555
         Visible         =   0   'False
         Width           =   345
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5145
         Left            =   60
         TabIndex        =   0
         Top             =   960
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   9075
         _Version        =   393216
         Cols            =   3
         RowHeightMin    =   20
         AllowBigSelection=   0   'False
         HighLight       =   0
         AllowUserResizing=   3
      End
      Begin Crystal.CrystalReport rptVoucher 
         Left            =   7680
         Top             =   525
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         Height          =   195
         Left            =   795
         TabIndex        =   6
         Top             =   225
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   930
         TabIndex        =   2
         Top             =   615
         Width           =   435
      End
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   120
      TabIndex        =   15
      Top             =   6345
      Width           =   2700
   End
End
Attribute VB_Name = "Frmpostingtogl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLmDoc As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Dim PR_VchType As New Recordset
Dim PR_Branch As New Recordset
Dim ls_sql As String
Const strChecked = "Y"
Const strUnChecked = "N"


Private Sub Check1_Click()

End Sub

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtuserid
    Set PO_DESC = txtuseriddesc
    Gs_SQL = "Select  userid 'User id' ,Username from syusers"
    Gs_FindFld = "Username"
    Gs_Subon = False
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by username"
    MyLookupOLDB.Caption = "Users"
    MyLookupOLDB.Show 1
End Sub
Private Sub Command1_Click()
If txtuserid <> "" Then
    Call Loadvouchers
Else
    Call MsgBox("Please enter userid ", vbCritical)
    txtuserid.SetFocus
End If
End Sub



Private Sub Command2_Click()
Call Postvoucher
End Sub
Private Sub Postvoucher()
            With Grid1
       
                For ln_cnt = 1 To .Rows - 1
                 
                 If .TextMatrix(ln_cnt, 8) = "P" Then
                    ls_sql = "Update Gl_ref set Pflag = 'P' where compcode = '" & Gs_compcode & "' and vchrtype = '" & .TextMatrix(ln_cnt, 3) & "'  and voucher_no = '" & .TextMatrix(ln_cnt, 4) & "' and value_date = '" & Format(DTPvoucherdate, "YYYY/MM/DD") & "'  "
                    gc_dbcon.Execute ls_sql
                    
                    ls_sql = "Update Gl_Trans set Pflag = 'P' where compcode = '" & Gs_compcode & "' and vchrtype = '" & .TextMatrix(ln_cnt, 3) & "'  and voucher_no = '" & .TextMatrix(ln_cnt, 4) & "' and value_date = '" & Format(DTPvoucherdate, "YYYY/MM/DD") & "'  "
                    gc_dbcon.Execute ls_sql
                 Else
                 
                    If Me.Caption = "GL Unposting" Then
                           If .TextMatrix(ln_cnt, 8) = "" Then
                               ls_sql = "Update Gl_ref set Pflag = Null where compcode = '" & Gs_compcode & "' and vchrtype = '" & .TextMatrix(ln_cnt, 3) & "'  and voucher_no = '" & .TextMatrix(ln_cnt, 4) & "' and value_date = '" & Format(DTPvoucherdate, "YYYY/MM/DD") & "'  "
                               gc_dbcon.Execute ls_sql
                       
                               ls_sql = "Update Gl_Trans set Pflag = Null  where compcode = '" & Gs_compcode & "' and vchrtype = '" & .TextMatrix(ln_cnt, 3) & "'  and voucher_no = '" & .TextMatrix(ln_cnt, 4) & "' and value_date = '" & Format(DTPvoucherdate, "YYYY/MM/DD") & "'  "
                               gc_dbcon.Execute ls_sql
                          End If
                           
                    End If
                 End If
                Next
            End With
            Call Loadvouchers
End Sub

Private Sub Form_Load()
 PR_VchType.Open "Select *,BranchCode+VchrType as Findfld from GlVchrType where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
 PR_Branch.Open "Select * From SysBranch where compcode = '" & Gs_compcode & "' Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
 

  DTPvoucherdate = Date
  InitializeGrid
End Sub
Public Sub InitializeGrid()
    With Grid1
        .Redraw = False
        .Clear
        .Rows = 2
        
        .RowHeight(1) = 320
        
        .FormatString = "Sr# |<Status|<Value Date|<Voucher Type|<Voucher No|<Narration|<Dr. Amount|<Cr. Amount|<chk"
        .ColWidth(1) = 700
        .ColWidth(2) = 1000
        .ColWidth(3) = 1200
        .ColWidth(4) = 1000
        .ColWidth(5) = 3000
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 0
        
        .Redraw = True
    End With
    PI_SrNo = 0
    PI_CurRow = 0
End Sub
Private Sub setprint()
'On Error GoTo LocalErr
Dim ls_BranchName As String
         If MySeek(Gs_BranchCode, "Branchcode", PR_Branch) Then ls_BranchName = PR_Branch("BranchDesc")
         If MySeek(txtvchrtype.Text, "VchrType", PR_VchType) Then ls_VchDesc = PR_VchType.Fields("Vchrdescrip")
        
   With rptVoucher
        
        .ReportFileName = App.Path & Gs_GlRepoPath & "\unpostVchr_Print.RPT"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & ls_VchDesc & "'"
        .Formulas(5) = "BranchName = '" & Gs_BranchCode + "-" + ls_BranchName & "'"
        .SelectionFormula = "{Gl_Trans.Voucher_No} = '" & Trim(txtvoucherno) & "' and {Gl_Trans.BranchCode} = '" & Gs_BranchCode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.VchrType} = '" & Trim(txtvchrtype) & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.CompCode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.Value_Date} = Date(" & Year(DTPvoucherdate) & "," & Month(DTPvoucherdate) & "," & Day(DTPvoucherdate) & ")"
        
        .Action = 1
   End With
 
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    PR_VchType.Close
    PR_Branch.Close
End Sub

Private Sub grid1_DblClick()
With Grid1
    txtvchrtype = .TextMatrix(.Row, 3)
    txtvoucherno = .TextMatrix(.Row, 4)
End With
Label2.Caption = "Processing Please Wait..."
Me.Refresh
 Call setprint
Label2.Caption = ""
Me.Refresh
End Sub

Private Sub Grid1_EnterCell()
   ' Call ProcessIndividualCells
   Grid1.CellBackColor = vbHighlight
End Sub
'Private Sub Grid1_KeyPress(KeyAscii As Integer)
'    With Grid1 ' Move the focus to next or prior cell
'        Select Case KeyAscii
'
'
'        Case vbKeyReturn, vbKeySpace
'            With Grid1
'                Call TriggerCheckbox(.Row, .Col)
'            End With
'
'        Case vbKeyReturn
'
'
'
'            'move to next cell.
'            If .Col + 1 <= .Cols - 1 Then
'                .Col = .Col + 1
'            Else
'                If .Row + 1 <= .Rows - 1 Then
'                    .Row = .Row + 1
'                    .Col = 0 + .FixedCols
'                Else
'                    .Row = 1
'                    .Col = 0 + .FixedCols
'                End If
'            End If
'        Case vbKeyTab
'            Select Case LeftOrRight
'            Case "Right"
'                If .Col + 1 <= .Cols - 1 Then
'                    .Col = .Col + 1
'                Else
'                    If .Row + 1 <= .Rows - 1 Then
'                        .Row = .Row + 1
'                        .Col = 0 + .FixedCols
'                    Else
'                        .Row = 1
'                        .Col = 0 + .FixedCols
'                    End If
'                End If
'            Case "Left"
'                If .Col > .FixedCols Then
'                    .Col = .Col - 1
'                Else
'                    If .Row > .FixedRows Then
'                        .Row = .Row - 1
'                        .Col = .Cols - 1
'                    Else
'                        .Row = .Rows - 1
'                        .Col = .Cols - 1
'                    End If
'                End If
'            End Select
'        Case vbKeyBack
'            'remove the last character, if any.
'            If Len(.Text) Then
'                .Text = Left(.Text, Len(.Text) - 1)
'            End If
'            Case Is < 32
'        Case Else
'            If .CellBackColor = vbHighlight Then
'                .Text = "": .CellBackColor = vbWindowBackground
'            End If
'            If Not .Col = 2 Then
'            .Text = .Text & Chr(KeyAscii)
'            End If
'        End Select
'    End With
'
'End Sub
Sub ProcessIndividualCells()
'    chkstatus.Visible = False   ' Make all controls invisible and turn on the current one
    
    With Grid1
        Select Case .Col
        Case 1 ' Cell the List1 is to show up in
            'See the following to make the Combo1 function
            'Sub List1_Click()
            'Sub List1_GotFocus()
            'Sub List1_KeyDown()
            'Sub List1_Validate()
            'Sub LoadCells()
            'Sub ProcessIndividualCells()
           ' chkstatus.Visible = True
           ' chkstatus.Move .CellLeft + .Left, .CellTop + .Top
           ' chkstatus.SetFocus
        
       Case 3 ' Cell the List1 is to show up in
            'See the following to make the Combo1 function
            'Sub List1_Click()
            'Sub List1_GotFocus()
            'Sub List1_KeyDown()
            'Sub List1_Validate()
            'Sub LoadCells()
            'Sub ProcessIndividualCells()
'             DtpTDate.Visible = True
 '            DtpTDate.Move .CellLeft + .Left, .CellTop + .Top
             'DtpTDate.SetFocus
        
        End Select
    End With
    
End Sub
Private Sub Grid1_LeaveCell()
Grid1.CellBackColor = vbWindowBackground
End Sub
Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
        With Grid1
'            If .TextMatrix(iRow, iCol) = strUnChecked Then
'                .TextMatrix(iRow, iCol) = strChecked
'                Set .CellPicture = picChecked.Picture
'            Else
'                .TextMatrix(iRow, iCol) = strUnChecked
'                Set .CellPicture = picUnchecked.Picture
'            End If
        If .Col = 1 Then
            If .CellPicture = picUnchecked.Picture Then
                Set .CellPicture = picChecked.Picture
               .TextMatrix(.Row, 8) = "P"
                
            Else
                Set .CellPicture = picUnchecked.Picture
                .TextMatrix(.Row, 8) = ""
            End If
        End If
        End With
End Sub
Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With Grid1
         If Not .Col = 2 Then
            If .MouseRow <> 0 And .MouseCol <> 0 Then
                Call TriggerCheckbox(.MouseRow, .MouseCol)
            End If
        End If
        End With
    End If
End Sub

Private Sub Loadvouchers()
Dim ln_cnt As Integer
Dim ls_opt As String
InitializeGrid
ln_cnt = 1
    
    ls_opt = IIf(Me.Caption = "GL Posting", " and Gl_Ref.Pflag is null", " and Gl_Ref.Pflag ='P'")
    
    ls_sql = "SELECT  Gl_Ref.Value_Date, Gl_Ref.Voucher_no, Gl_Ref.VchrType, Gl_Ref.Vchr_Remarks, SUM(Gl_Trans.Dr_Amount) AS Dr_amount, SUM(Gl_Trans.Cr_Amount)"
    ls_sql = ls_sql & " AS Cr_amount FROM  Gl_Ref INNER JOIN Gl_Trans ON Gl_Ref.CompCode = Gl_Trans.compcode AND Gl_Ref.BranchCode = Gl_Trans.BranchCode AND"
    ls_sql = ls_sql & " Gl_Ref.Value_Date = Gl_Trans.Value_Date And Gl_Ref.Voucher_No = Gl_Trans.Voucher_No And Gl_Ref.VchrType = Gl_Trans.VchrType"
    ls_sql = ls_sql & " where Gl_Ref.compcode = '" & Gs_compcode & "' and Gl_Ref.Value_Date = '" & Format(DTPvoucherdate, "YYYY/MM/DD") & "' and Gl_Ref.UserId = '" & txtuserid & "'" & ls_opt
    ls_sql = ls_sql & " GROUP BY Gl_Ref.Value_Date, Gl_Ref.Voucher_no, Gl_Ref.VchrType, Gl_Ref.Vchr_Remarks"

   
   pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
   If Not pr_dumy.EOF Then
        With Grid1
                


            While Not pr_dumy.EOF
                .Row = .Rows - 1
                
                .TextMatrix(.Row, 0) = ln_cnt
                .Col = 1
                .Row = .Row
                 Set .CellPicture = IIf(chkselectall.Value = 1, picChecked, picUnchecked)
                .Col = 2
                .Row = .Row
                .TextMatrix(.Row, 2) = pr_dumy("Value_date")
                .Col = 3
                .Row = .Row
                .TextMatrix(.Row, 3) = pr_dumy("vchrtype")
                .TextMatrix(.Row, 4) = pr_dumy("voucher_no")
                .TextMatrix(.Row, 5) = pr_dumy("vchr_remarks")
                .TextMatrix(.Row, 6) = pr_dumy("Dr_Amount")
                .TextMatrix(.Row, 7) = pr_dumy("Cr_Amount")
                .TextMatrix(.Row, 8) = IIf(chkselectall.Value = 1, "P", "")
                
                .Rows = .Rows + 1
                 ln_cnt = ln_cnt + 1
                 pr_dumy.MoveNext
                
             Wend
           
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        Else
        txtuserid.SetFocus
        ProcessIndividualCells
        Call MsgBox(Gs_RecNFMsg, vbCritical)
        End If
pr_dumy.Close
End Sub

Private Sub txtuserid_Change()
txtuseriddesc = ""
End Sub

Private Sub txtuserid_Validate(Cancel As Boolean)
If txtuserid <> "" Then
    pr_dumy.Open "Select * from syusers where compcode = '" & Gs_compcode & "' and userid = '" & txtuserid & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    
    If Not pr_dumy.EOF Then
        txtuseriddesc = pr_dumy("username")
        DTPvoucherdate.SetFocus
    Else
        Call MsgBox("Record not found", vbCritical)
            txtuserid.SetFocus
            txtuseriddesc = ""
    End If
    pr_dumy.Close
End If
End Sub
