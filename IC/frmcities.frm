VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcities 
   Caption         =   "Cities Setup"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmcities.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6405
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
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   6360
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   3060
         Picture         =   "frmcities.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1410
         Width           =   315
      End
      Begin VB.TextBox txtaccount2 
         Height          =   315
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1425
         Width           =   1365
      End
      Begin VB.TextBox txtaccountdesc2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3405
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1410
         Width           =   2865
      End
      Begin VB.CommandButton CmdAccount1 
         Height          =   315
         Left            =   3060
         Picture         =   "frmcities.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   990
         Width           =   315
      End
      Begin VB.TextBox txtaccount1 
         Height          =   315
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1005
         Width           =   1365
      End
      Begin VB.TextBox txtaccountdesc1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3405
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   990
         Width           =   2865
      End
      Begin VB.TextBox txtcities 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1665
         MaxLength       =   3
         TabIndex        =   7
         Tag             =   "SKIP"
         Top             =   255
         Width           =   480
      End
      Begin VB.TextBox Textx 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   4620
         MaxLength       =   35
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1665
         MaxLength       =   25
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   615
         Width           =   4605
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
         Left            =   2160
         Picture         =   "frmcities.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "GL Code For Client Sub Ledger :"
         Height          =   495
         Left            =   60
         TabIndex        =   9
         Top             =   1365
         Width           =   1530
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "GL Code For Vendor Sub Ledger :"
         Height          =   495
         Left            =   60
         TabIndex        =   8
         Top             =   915
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   510
         TabIndex        =   6
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   705
         TabIndex        =   5
         Top             =   615
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
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
               Picture         =   "frmcities.frx":0760
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcities.frx":0BB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcities.frx":1008
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcities.frx":145C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcities.frx":18B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcities.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcities.frx":2458
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmcities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Cities As New Recordset
Dim pr_dumy As New Recordset

Private Sub CmdAccount1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount1
    Set PO_DESC = txtaccountdesc1
    Gs_SQL = "Select  Acct_sub1+Acct_sub2 'Account Code' ,Acct_Desc as Description from Gl_sub2"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount1) > 0 Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcities
    Set PO_DESC = txtdesc
    Gs_SQL = "Select Citycode, cityname from Cities "
    Gs_FindFld = "Cityname"
    Gs_OrderBy = "Order by cityname"
    
    MyLookupOLDB.Caption = "Cities"
    MyLookupOLDB.Show 1
    
    If Len(txtcities) > 0 Then Txtcities_KeyDown vbKeyReturn, vbKeyShift

End Sub
Private Function maxtranscode() As String
Dim pr_dumy As New Recordset
pr_dumy.Open "select max(cityCode) as transcode from cities", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
pr_dumy.Close
End Function

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount2
    Set PO_DESC = txtaccountdesc2
    Gs_SQL = "Select  Acct_sub1+Acct_sub2 'Account Code' ,Acct_Desc as Description from Gl_sub2"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount2) > 0 Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Cities, Me, txtcities, txtdesc, "X", "CompCount", 3, "CityCode", "CityName", 1, False, Toolbar1)
End Sub
Private Sub txtaccount1_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount1 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select Acct_sub1+Acct_sub2 'Account Code' ,Acct_Desc as Description from Gl_sub2 where compcode = '" & Gs_compcode & "' and  Acct_sub1+Acct_sub2 = '" & txtaccount1 & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                txtaccountdesc1 = pr_dumy("description")
                txtaccount2.SetFocus
            End If
         pr_dumy.Close

End If

End Sub
Private Sub txtaccount2_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount2 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select Acct_sub1+Acct_sub2 'Account Code' ,Acct_Desc as Description from Gl_sub2 where compcode = '" & Gs_compcode & "' and  Acct_sub1+Acct_sub2 = '" & txtaccount2 & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                txtaccountdesc2 = pr_dumy("description")
                txtdesc.SetFocus
            End If
         pr_dumy.Close

End If

End Sub
Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SRCIT00001")
  SetToolBar(2) = chkRights("SRCIT00002")
  SetToolBar(3) = chkRights("SRCIT00003")
  SetToolBar(4) = chkRights("SRCIT00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
'  PR_GlDetail.Open "Select AccountNo,Acct_Desc FROM GL_detail Where Acct_Type ='B' and compcode = '" & Gs_compcode & "' Order By AccountNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_Cities.Open "Select * from Cities Order By CityCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
 ' PR_VchType.Open "Select * from GlVchrType where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  
  PB_BlnkLoca = IIf(PR_Cities.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Cities.Close
End Sub


Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtaccount1.SetFocus
End Sub

Private Sub Txtcities_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtcities.Text = DoPad(txtcities.Text, txtcities.MaxLength)
      lb_found = MySeek(txtcities.Text, "CityCode", PR_Cities)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtcities.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtcities.SetFocus
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
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
    Else
      cmdLookup.Enabled = True
    End If
    
  If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Cities, Me, txtcities, txtdesc, "X", "CompCount", 3, "CityCode", "CityName", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
     txtcities.Enabled = False
     txtcities = maxtranscode
     txtdesc.SetFocus
    Else
     txtcities.Enabled = True
    End If
End Sub

Public Sub SaveValues()
PB_BlnkLoca = False
'Dim ls_btype As String

'ls_btype = IIf(chklocal, "L", "F")
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT into Cities(CityCode,CityName,GLVendor,GLclient) VALUES ('" & txtcities.Text & "','" & txtdesc.Text & "','" & txtaccount1 & "','" & txtaccount2 & "')"
           Case "E"
              gc_dbcon.Execute "UPDATE Cities SET CityName= '" & txtdesc.Text & "',GLVendor = '" & txtaccount1 & "',GLClient ='" & txtaccount2 & "' WHERE  CityCode= '" & txtcities.Text & "'"
           Case "D"
              gc_dbcon.Execute "DELETE FROM Cities WHERE CityCode = '" & txtcities.Text & "'"
     End Select
PR_Cities.Requery
    If Mode = "A" Then
     txtcities.Enabled = False
     txtcities = maxtranscode
     txtdesc.SetFocus
    Else
     txtcities.Enabled = True
    End If
End Sub

Private Sub SetVal()
     txtdesc = Trim(PR_Cities("CityName") & "")
     txtaccount1 = Trim(PR_Cities("GLVendor") & "")
     If txtaccount1 <> "" Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
     txtaccount2 = Trim(PR_Cities("GLClient") & "")
     If txtaccount2 <> "" Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtcities.Text) = txtcities.MaxLength And txtdesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  PR_Cities.Requery
End Sub

