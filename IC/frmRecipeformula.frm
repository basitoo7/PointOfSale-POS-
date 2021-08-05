VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRecipeformula 
   Caption         =   "Item Recipe formula"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecipeformula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtitemcost 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   9810
      Locked          =   -1  'True
      MaxLength       =   64
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "SKIP"
      Top             =   6450
      Width           =   1380
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
      Height          =   5835
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   11310
      Begin VB.TextBox txtamount 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5550
         MaxLength       =   10
         TabIndex        =   15
         Top             =   960
         Width           =   1425
      End
      Begin VB.TextBox txtavgrate 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3285
         MaxLength       =   10
         TabIndex        =   13
         Top             =   945
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add To Grid"
         Height          =   330
         Left            =   10125
         TabIndex        =   12
         Top             =   945
         Width           =   1065
      End
      Begin VB.TextBox txtqty 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2055
         Picture         =   "frmRecipeformula.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   555
         Width           =   315
      End
      Begin VB.TextBox txtrecipecode 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   7
         Top             =   570
         Width           =   690
      End
      Begin VB.TextBox txtrecipedesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2415
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   555
         Width           =   8820
      End
      Begin VB.TextBox txtitemdesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2415
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   165
         Width           =   8805
      End
      Begin VB.TextBox txtitemcode 
         BackColor       =   &H00FFFF80&
         Height          =   330
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "SKIP"
         Top             =   180
         Width           =   720
      End
      Begin VB.CommandButton cmdLookup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2070
         Picture         =   "frmRecipeformula.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   165
         Width           =   315
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   4290
         Left            =   90
         TabIndex        =   17
         Top             =   1440
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   7567
         _Version        =   393216
         RowHeightMin    =   300
         BackColorSel    =   16777215
         ForeColorSel    =   0
         GridColor       =   8421504
         AllowBigSelection=   0   'False
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Amount :"
         Height          =   210
         Left            =   4755
         TabIndex        =   16
         Top             =   990
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Avg Rate :"
         Height          =   210
         Left            =   2565
         TabIndex        =   14
         Top             =   975
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quantity :"
         Height          =   210
         Left            =   495
         TabIndex        =   10
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Recipe Code :"
         Height          =   210
         Left            =   195
         TabIndex        =   9
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item Code :"
         Height          =   210
         Left            =   405
         TabIndex        =   2
         Top             =   195
         Width           =   795
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
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
            Caption         =   "&Listing"
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
               Picture         =   "frmRecipeformula.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipeformula.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipeformula.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipeformula.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipeformula.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipeformula.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipeformula.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Item Cost :"
      Height          =   255
      Left            =   8925
      TabIndex        =   19
      ToolTipText     =   "Enter Value Date"
      Top             =   6495
      Width           =   870
   End
End
Attribute VB_Name = "frmRecipeformula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Dumy As New Recordset
Dim PR_Dumy1 As New Recordset
Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String
Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtItemCode
    Set PO_DESC = txtItemDesc
    Gs_SQL = "Select itemCode, Description from IC_Item "
    Gs_FindFld = "Description"
    Gs_OtherPara = " Where compcode = '" & Gs_compcode & "' and mtcode ='002'"
    Gs_OrderBy = "Order by Description"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
    
    If txtItemCode <> "" Then Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub TotalAmount()

    Dim ln_cnt As Integer
      txtitemcost = ""
  
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txtitemcost = Val(txtitemcost) + Val(.TextMatrix(ln_cnt, 6))
        Next
    End With
    
  
End Sub

Private Sub Command2_Click()
If Trim(txtItemCode) = "" Then
    Call MsgBox("Enter Item Code !!!", vbCritical)
    txtItemCode.SetFocus
ElseIf Trim(txtrecipecode) = "" Then
    Call MsgBox("Enter Recipe Code !!!", vbCritical)
    txtrecipecode.SetFocus
ElseIf Trim(txtqty) = "" Then
    Call MsgBox("Enter Quantity !!!", vbCritical)
    txtqty.SetFocus
ElseIf Trim(txtavgrate) = "" Then
    Call MsgBox("Enter Avg Rate !!!", vbCritical)
    txtavgrate.SetFocus
ElseIf SearchInGrid(GrdGRN, txtrecipecode, 2) Then
    Call MsgBox("Item already exist in Grid !!!", vbCritical)
    txtrecipecode.SetFocus

Else
    Call AddToGrid
End If

End Sub

Private Sub TxtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtamount <> "" Then
   Command2_Click
End If
End Sub

Private Sub txtavgrate_LostFocus()
If txtavgrate <> "" Then
txtamount = Val(txtqty) * Val(txtavgrate)
End If

End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtItemCode) <> "" And KeyCode = vbKeyReturn Then
        txtItemCode = DoPad(txtItemCode, txtItemCode.MaxLength)
        PR_Dumy.Open "Select * from IC_Item where itemcode = '" & txtItemCode & "' and compcode = '" & Gs_compcode & "' and mtcode = '002'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Item not found !!!", vbCritical)
            txtItemCode = ""
            txtItemDesc = ""
            txtItemCode.SetFocus
        Else
            
                txtItemDesc = PR_Dumy("Description")
                If txtrecipecode.Enabled Then txtrecipecode.SetFocus
            If Mode <> "A" Then
                Call SetVal
            Else
            PR_Dumy1.Open "Select * from IC_RecipeFormula where itemcode = '" & txtItemCode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If Not PR_Dumy1.EOF Then
                Call MsgBox("Item already exist !!!", vbCritical)
                txtItemCode = ""
                txtItemDesc = ""
                txtItemCode.SetFocus
                InitializeGrid
            End If
            PR_Dumy1.Close
            
            End If
        End If
        PR_Dumy.Close

ElseIf Trim(txtItemCode) = "" And KeyCode = vbKeyReturn Then
        txtItemCode = ""
        txtItemDesc = ""
End If
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtrecipecode
    Set PO_DESC = txtrecipedesc
    Gs_SQL = "Select itemCode, Description from IC_Item "
    Gs_FindFld = "Description"
    Gs_OtherPara = " Where compcode = '" & Gs_compcode & "' and mtcode <>'002'"
    Gs_OrderBy = "Order by Description"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
   
    
    If txtrecipecode <> "" Then Call txtrecipecode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtqty_LostFocus()
If txtqty <> "" Then
txtamount = Val(txtqty) * Val(txtavgrate)
End If
End Sub

Private Sub txtrecipecode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtrecipecode) <> "" And KeyCode = vbKeyReturn Then
        txtrecipecode = DoPad(txtrecipecode, txtrecipecode.MaxLength)
        PR_Dumy.Open "Select * from IC_item where itemcode = '" & txtrecipecode & "' and compcode = '" & Gs_compcode & "' and mtcode <> '002' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Recipe code not found !!!", vbCritical)
            txtrecipecode = ""
            txtrecipedesc = ""
            txtrecipecode.SetFocus
        Else
            txtrecipedesc = Trim(PR_Dumy("Description") & "")
            txtavgrate = Val(0 & PR_Dumy("avgrate"))
            txtqty.SetFocus
        End If
        PR_Dumy.Close

ElseIf Trim(txtrecipecode) = "" And KeyCode = vbKeyReturn Then
        txtrecipecode = ""
        txtrecipedesc = ""
End If
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("ICRECFOR01")
  SetToolBar(2) = chkRights("ICRECFOR02")
  SetToolBar(3) = chkRights("ICRECFOR03")
  SetToolBar(4) = chkRights("ICRECFOR04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  InitializeGrid

End Sub


Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call GrdGRN_DblClick
   If KeyCode = vbKeyDelete Then
       With GrdGRN
          If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
             .RemoveItem .Row
             If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
             End If
       End With
   End If
End Sub


Private Sub GrdGRN_DblClick()
  With GrdGRN
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
       txtrecipecode = .TextMatrix(.Row, 2)
       txtrecipedesc = .TextMatrix(.Row, 3)
       txtqty = .TextMatrix(.Row, 4)
       txtavgrate = .TextMatrix(.Row, 5)
       txtamount = .TextMatrix(.Row, 6)
       
       PS_RowClicked = "Y"
       
       txtrecipecode.SetFocus
    End With
End Sub
Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtqty <> "" Then
   txtavgrate.SetFocus
End If
End Sub
Private Sub txtavgrate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtavgrate <> "" Then
   txtamount.SetFocus
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Sites, Me, txtItemCode, txtItemDesc, "X", "CompCount", 3, "SiteCode", "Description", 1, False, Toolbar1)
    End If
    If Button.Index = 7 Then InitializeGrid
    
    If Button.Index = 1 And Mode = "A" Then InitializeGrid
    
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
                                   .TextMatrix(.Row, 1) = Trim(txtItemCode)
                                   .TextMatrix(.Row, 2) = Trim(txtrecipecode)
                                   .TextMatrix(.Row, 3) = Trim(txtrecipedesc)
                                   .TextMatrix(.Row, 4) = Trim(txtqty)
                                   .TextMatrix(.Row, 5) = Trim(txtavgrate)
                                   .TextMatrix(.Row, 6) = Trim(txtamount)
                                                
                                txtqty = ""
                                txtrecipecode = ""
                                txtrecipedesc = ""
                                txtavgrate = ""
                                txtamount = ""
                                TotalAmount
                                
                                
                                
                                PS_RowClicked = ""
                        End With
                    
                        
                        txtrecipecode.SetFocus
End Sub

Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Item Code|<Recipe Code|<Description|<Quantity|<AvgRate|<Amount"
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 3000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        
        .Redraw = True
    End With
    PI_SrNo = 0
    PI_CurRow = 0
    PS_RowClicked = ""
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim ls_sql As String
PB_BlnkLoca = False
gc_dbcon.BeginTrans
     Select Case Mode
           Case "D"
              ls_sql = "DELETE FROM IC_RecipeFormula WHERE ItemCode = '" & txtItemCode.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
            
           Case Else
              ls_sql = "DELETE FROM IC_RecipeFormula WHERE ItemCode = '" & txtItemCode.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
            
              With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      ls_sql = "INSERT into IC_RecipeFormula(compcode,ItemCode,RecipeCode,Quantity,Rate,Amount) VALUES ('" & Gs_compcode & "','" & .TextMatrix(ln_cnt, 1) & "','" & .TextMatrix(ln_cnt, 2) & "'," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & "," & Val(.TextMatrix(ln_cnt, 6)) & ")"
                      gc_dbcon.Execute ls_sql
                    Next
              End With
              
              
     End Select

gc_dbcon.CommitTrans
InitializeGrid
txtItemCode.Text = ""
txtItemDesc.Text = ""
txtItemCode.SetFocus

Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Public Sub ClearVal()
     txtLocation = ""
     txtDesc = ""
End Sub

Private Sub SetVal()
On Error GoTo LocalErr

Dim pr_dumyloadtrans As New Recordset
Dim PR_Dumy1 As New Recordset

InitializeGrid
    
    pr_dumyloadtrans.Open "select * from IC_RecipeFormula where compcode = '" & Gs_compcode & "' and ItemCode = '" & txtItemCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly
   
    If Not pr_dumyloadtrans.EOF Then
        With GrdGRN
            Do While Not pr_dumyloadtrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                
                .TextMatrix(.Row, 1) = Trim(pr_dumyloadtrans("Itemcode") & "")
                .TextMatrix(.Row, 2) = Trim(pr_dumyloadtrans("RecipeCode") & "")
                 PR_Dumy1.Open "Select * from IC_item where itemcode = '" & .TextMatrix(.Row, 2) & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                   If Not PR_Dumy1.EOF Then
                    .TextMatrix(.Row, 3) = Trim(PR_Dumy1("Description") & "")
                   End If
                 PR_Dumy1.Close
                
                .TextMatrix(.Row, 4) = Val(0 & pr_dumyloadtrans("Quantity"))
                .TextMatrix(.Row, 5) = Val(0 & pr_dumyloadtrans("Rate"))
                .TextMatrix(.Row, 6) = Val(0 & pr_dumyloadtrans("Amount"))
                .Rows = .Rows + 1
                pr_dumyloadtrans.MoveNext
                If pr_dumyloadtrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
            PI_SrNo = .Rows - 1
        End With
    End If
    pr_dumyloadtrans.Close
    TotalAmount
Exit Sub
LocalErr:
Call MsgBox(Err.Description)
     
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtItemCode.Text) = txtItemCode.MaxLength And PI_SrNo > 0 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtDesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub

