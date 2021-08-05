VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecipes 
   Caption         =   "Recipes Setup"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecipes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1785
      Left            =   30
      TabIndex        =   1
      Top             =   615
      Width           =   4935
      Begin VB.TextBox txtrecipeRate 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   975
         Width           =   1005
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
         Left            =   1770
         Picture         =   "frmRecipes.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1365
         Width           =   315
      End
      Begin VB.TextBox txtmdesc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2115
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1365
         Width           =   2700
      End
      Begin VB.TextBox txtmcode 
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1365
         Width           =   435
      End
      Begin VB.TextBox Textx 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   4620
         MaxLength       =   35
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   3495
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
         Left            =   1815
         Picture         =   "frmRecipes.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtRecipeCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Recipe Rate :"
         Height          =   210
         Left            =   300
         TabIndex        =   13
         Top             =   1005
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "M-Unit :"
         Height          =   210
         Left            =   720
         TabIndex        =   11
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   375
         TabIndex        =   6
         Top             =   600
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
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
            Caption         =   "Refresh"
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
               Picture         =   "frmRecipes.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipes.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipes.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipes.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipes.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipes.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecipes.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmRecipes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Recipe As New Recordset
Dim pr_dumy As New Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtRecipeCode
    Set PO_DESC = txtdesc
    
    GoTop PR_Recipe
    MyLookup.Caption = "Recipes"
    MyLookup.FillGrid PR_Recipe, "RecipeCode", "RecipeDesc", txtRecipeCode.MaxLength
    MyLookup.Show 1
    
    If Len(txtRecipeCode) > 0 Then txtrecipecode_KeyDown vbKeyReturn, vbKeyShift

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Recipe, Me, txtRecipeCode, txtdesc, "X", "CompCount", 3, "RecipeCode", "RecipeDesc", 1, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("CITY000001")
  SetToolBar(2) = chkRights("CITY000002")
  SetToolBar(3) = chkRights("CITY000003")
  SetToolBar(4) = chkRights("CITY000004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  

  PR_Recipe.Open "Select * from IC_Recipes Order By RecipeCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  PB_BlnkLoca = IIf(PR_Recipe.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Recipe.Close
End Sub


Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtrecipeRate.SetFocus
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtmcode
    Set PO_DESC = txtmdesc
    Gs_SQL = "Select MCode, Description from IC_ItemUM "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    MyLookupOLDB.Caption = "Item Units"
    MyLookupOLDB.Show 1
    
    If txtmcode <> "" Then Call txtmcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtmcode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtmcode) <> "" And KeyCode = vbKeyReturn Then
        txtmcode = DoPad(txtmcode, 3)
        pr_dumy.Open "Select * from IC_ItemUM where Mcode = '" & txtmcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Item Unit not found !!!", vbCritical)
            txtmcode = ""
            txtmdesc = ""
            txtmcode.SetFocus
        Else
            txtmdesc = pr_dumy("Description")
        End If
        pr_dumy.Close

ElseIf Trim(txtmcode) = "" And KeyCode = vbKeyReturn Then
        txtmcode = ""
        txtmdesc = ""
End If

End Sub

Private Sub txtrecipecode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtRecipeCode.Text = DoPad(txtRecipeCode.Text, txtRecipeCode.MaxLength)
      lb_found = MySeek(txtRecipeCode.Text, "RecipeCode", PR_Recipe)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtRecipeCode.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtRecipeCode.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                     ' txtLocation.Enabled = False
                      txtdesc.SetFocus
                   End If
                End If
            End Select
ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
  If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
  Else
      Mode = DentMode(Mode, Button.Index, PR_Recipe, Me, txtRecipeCode, txtdesc, "X", "CompCount", 3, "RecipeCode", "RecipeDesc", 1, False, Toolbar1)
  End If
  If Mode = "A" Then
      cmdLookup.Enabled = False
      txtRecipeCode.Enabled = False
      txtRecipeCode = maxtranscode
      txtdesc.SetFocus
  Else
      cmdLookup.Enabled = True
      txtRecipeCode.Enabled = True
      txtRecipeCode.SetFocus
  End If
      
    
    
End Sub

Public Sub SaveValues()
PB_BlnkLoca = False
'Dim ls_btype As String

'ls_btype = IIf(chklocal, "L", "F")
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT into IC_Recipes(RecipeCode,RecipeDesc,RecipeRate,Mcode) VALUES ('" & txtRecipeCode.Text & "','" & txtdesc.Text & "'," & Val(txtrecipeRate) & ",'" & txtmcode.Text & "')"
           Case "E"
              gc_dbcon.Execute "UPDATE IC_Recipes SET RecipeDesc= '" & txtdesc.Text & "',RecipeRate = " & Val(txtrecipeRate) & " ,mcode = '" & txtmcode & "'  WHERE  RecipeCode= '" & txtRecipeCode.Text & "'"
           Case "D"
              gc_dbcon.Execute "DELETE FROM IC_Recipes WHERE RecipeCode = '" & txtRecipeCode.Text & "'"
     End Select
PR_Recipe.Requery
If Mode = "A" Then
txtRecipeCode = maxtranscode
End If
End Sub

Private Sub SetVal()
     txtdesc = Trim(PR_Recipe("RecipeDesc") & "")
     txtrecipeRate = Val(PR_Recipe("RecipeRate"))
     txtmcode = Trim(PR_Recipe("mcode") & "")
     If txtmcode <> "" Then Call txtmcode_KeyDown(vbKeyReturn, vbKeyShift)
     
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtRecipeCode.Text) = txtRecipeCode.MaxLength And txtdesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  PR_Recipe.Requery
End Sub
Private Function maxtranscode() As String
Dim pr_dumy As New Recordset
pr_dumy.Open "select max(RecipeCode) as transcode from IC_Recipes", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
pr_dumy.Close
End Function

Private Sub txtrecipeRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtmcode.SetFocus
End If
End Sub
