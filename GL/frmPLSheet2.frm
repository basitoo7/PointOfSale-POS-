VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPLSheet2 
   Caption         =   "Profit and Loss Notes Setup"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   Icon            =   "frmPLSheet2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   15
      TabIndex        =   1
      Top             =   570
      Width           =   4935
      Begin VB.CommandButton cmdLookup0 
         Height          =   315
         Left            =   1815
         Picture         =   "frmPLSheet2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup1 
         Height          =   315
         Left            =   1920
         Picture         =   "frmPLSheet2.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   315
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1020
         Width           =   3495
      End
      Begin VB.TextBox txtGlDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2160
         MaxLength       =   64
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   225
         Width           =   2655
      End
      Begin VB.TextBox txtsub0 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtDetailCode 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1305
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "SKIP"
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   195
         Left            =   390
         TabIndex        =   8
         Top             =   1050
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PL Code :"
         Height          =   195
         Left            =   555
         TabIndex        =   7
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Note Code :"
         Height          =   195
         Left            =   405
         TabIndex        =   6
         Top             =   645
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      _ExtentX        =   8811
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
               Picture         =   "frmPLSheet2.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPLSheet2.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPLSheet2.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPLSheet2.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPLSheet2.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPLSheet2.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPLSheet2.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPLSheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_GlPL01 As Recordset
Dim PR_GlPLN As Recordset
Dim pr_dumy As New Recordset

Private Sub cmdLookup0_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsub0
    Set PO_DESC = txtglDesc
    GoTop PR_GlPL01
    MyLookup.Caption = "Profit Loss MAIN HEADS"
    MyLookup.FillGrid PR_GlPL01, "PLCODE", "PLDESC", 3
    MyLookup.Show 1
    
    If Len(txtsub0) > 0 Then
        txtsub0_Validate False
        SendKeys vbTab
    End If

End Sub

Private Sub cmdLookup1_Click()
    
    If Trim(txtsub0.Text) = "" Then
        MsgBox "Select Profit and Loss Code!!!", vbInformation, App.ProductName
        txtsub0.SetFocus
        Exit Sub
    End If
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtDetailCode
    Set PO_DESC = txtdesc
    GoTop PR_GlPLN
    PR_GlPLN.Filter = "PLCODE = '" & txtsub0 & "'"
    MyLookup.Caption = "Profit and Loss Notes"
    MyLookup.FillGrid PR_GlPLN, "PLNCODE", "PLNDESC", 6
    MyLookup.Show 1
    PR_GlPLN.Filter = adFilterNone
    
    
    If Len(txtDetailCode) > 0 Then
        txtDetailCode_Validate False
        SendKeys vbTab
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("GLFRM20001")
  SetToolBar(2) = chkRights("GLFRM20002")
  SetToolBar(3) = chkRights("GLFRM20003")
  SetToolBar(4) = chkRights("GLFRM20004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  
  Set PR_GlPL01 = New Recordset
  Set PR_GlPLN = New Recordset
   
  PR_GlPL01.Open "Select GL_PlSheet1.* from GL_PlSheet1 where compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_GlPLN.Open "Select GL_PlSheet2.* from GL_PlSheet2 where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
   
  PB_BlnkGls0 = IIf(PR_GlPLN.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GlPLN.Close
    PR_GlPL01.Close
End Sub

Private Sub txtDetailCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdLookup1_Click
    End If
End Sub

Private Sub txtDetailCode_Validate(Cancel As Boolean)
Dim lb_found As Boolean

    If txtDetailCode <> "" And Mode <> "A" Then
        PR_GlPLN.Filter = "PLCODE = '" & txtsub0 & "'"
        lb_found = MySeek(Trim(txtDetailCode.Text), "PLNCODE", PR_GlPLN)
        Select Case Mode
            Case "A"
                If lb_found Then
                    Call SetErr(Gs_RecFdMsg, vbCritical)
                    'Cancel = True
                    SetClear Me
                End If
            Case Else
                If Not lb_found Then
                    Call SetErr(Gs_RecNFMsg, vbCritical)
                    'Cancel = True
                    SetClear Me
                Else
                    Call SetVal
                End If
        End Select
     PR_GlPLN.Filter = adFilterNone
    Else
        txtDetailCode = ""
        txtdesc = ""
    End If
End Sub

Private Sub txtsub0_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdLookup0_Click
    End If
End Sub
  
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_BlnkGls0 And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlPL01, Me, txtsub0, txtdesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
       If Mode = "A" Then
            cmdLookup1.Enabled = False
      Else
           cmdLookup0.Enabled = True
            cmdLookup1.Enabled = True
            txtDetailCode.Enabled = True
       End If
    End If
End Sub

Public Function ChkInputs() As Boolean
    If Len(txtDetailCode.Text) = 4 And Len(RTrim(txtdesc)) > 0 And Len(txtsub0.Text) = 3 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
   PR_GlPL01.Requery
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim ls_accNature As String

PB_BlnkGls0 = False
gc_dbcon.BeginTrans
        If Mode = "A" Then
                pr_dumy.Open "select max(PLNCODE) as PLNCODE from GL_PlSheet2 where PLCODE = '" & txtsub0.Text & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                If Not pr_dumy.EOF Then
                    txtDetailCode = DoPad(Trim(str(Val(0 & pr_dumy("PLNCODE")) + 1)), txtDetailCode.MaxLength)
                Else
                    txtDetailCode = DoPad(Trim(str(1)), txtDetailCode.MaxLength)
                End If
                pr_dumy.Close
        End If


     Select Case Mode
           Case "A"
             gc_dbcon.Execute "INSERT INTO GL_PlSheet2(Compcode,PLCODE, PLNCODE,PLNDESC) VALUES ('" & Gs_compcode & "',  '" & txtsub0 & "','" & txtDetailCode & "','" & UCase(RepApp(Trim(txtdesc.Text))) & "')"
              
           Case "E"
              gc_dbcon.Execute "UPDATE GL_PlSheet2 SET PLNDESC = '" & UCase(RepApp(Trim(txtdesc.Text))) & "' WHERE GL_PlSheet2.PLNCODE = '" & txtDetailCode.Text & "' AND GL_PlSheet2.PLCODE = '" & txtsub0.Text & "' and compcode = '" & Gs_compcode & "'"
              txtDetailCode.Enabled = True
              
           Case "D"
              gc_dbcon.Execute "DELETE FROM GL_PlSheet2 WHERE GL_PlSheet2.PLNCODE = '" & txtDetailCode.Text & "' AND GL_PlSheet2.PLCODE = '" & txtsub0.Text & "' and compcode = '" & Gs_compcode & "' "
              
     End Select
gc_dbcon.CommitTrans
PR_GlPLN.Requery

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub SetVal()
     txtsub0 = Trim(PR_GlPLN("PLCODE"))
     txtsub0_Validate True
     txtdesc = Trim(PR_GlPLN("PLNDESC"))
     
End Sub

Private Sub txtsub0_Validate(Cancel As Boolean)
Dim lb_found As Boolean

    If txtsub0.Text <> "" Then
        txtsub0.Text = DoPad(UCase(txtsub0.Text), txtsub0.MaxLength)
        lb_found = MySeek(txtsub0.Text, "PLCODE", PR_GlPL01)
       
        If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            'Cancel = True
            txtsub0.Text = ""
            txtglDesc.Text = ""
            SetClear Me
        Else
            txtglDesc = PR_GlPL01("PLDESC")
            PR_GlPLN.Filter = adFilterNone
            PR_GlPLN.Filter = "PLCODE = '" & txtsub0 & "'"
            If Mode = "A" Then
                pr_dumy.Open "select max(PLNCODE) as PLNCODE from GL_PlSheet2 where PLCODE = '" & txtsub0.Text & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                If Not pr_dumy.EOF Then
                    txtDetailCode = DoPad(Trim(str(Val(0 & pr_dumy("PLNCODE")) + 1)), txtDetailCode.MaxLength)
                Else
                    txtDetailCode = DoPad(Trim(str(1)), txtDetailCode.MaxLength)
                End If
                pr_dumy.Close
             End If
        End If
    End If
    
End Sub
