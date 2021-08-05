VERSION 5.00
Begin VB.Form frmPOPurchaseAccountSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Accounts Setup"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   Icon            =   "PurchaseAccountSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3585
      MaxLength       =   50
      TabIndex        =   46
      Top             =   3960
      Width           =   2865
   End
   Begin VB.TextBox txtAccount11 
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
      Left            =   1830
      MaxLength       =   50
      TabIndex        =   45
      Top             =   3990
      Width           =   1365
   End
   Begin VB.CommandButton CMDAccount11 
      Height          =   315
      Left            =   3210
      Picture         =   "PurchaseAccountSetup.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3975
      Width           =   315
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000080&
      Height          =   5355
      Left            =   15
      TabIndex        =   1
      Top             =   -60
      Width           =   6540
      Begin VB.CommandButton CMDAccount10 
         Height          =   315
         Left            =   3210
         Picture         =   "PurchaseAccountSetup.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3645
         Width           =   315
      End
      Begin VB.TextBox txtAccount10 
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
         Left            =   1815
         MaxLength       =   50
         TabIndex        =   41
         Top             =   3675
         Width           =   1365
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   40
         Top             =   3645
         Width           =   2865
      End
      Begin VB.CommandButton CMDAccount9 
         Height          =   315
         Left            =   3225
         Picture         =   "PurchaseAccountSetup.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3300
         Width           =   315
      End
      Begin VB.TextBox txtAccount9 
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
         Left            =   1815
         MaxLength       =   50
         TabIndex        =   37
         Top             =   3315
         Width           =   1365
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   36
         Top             =   3285
         Width           =   2865
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   34
         Top             =   2910
         Width           =   2865
      End
      Begin VB.TextBox txtAccount8 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2940
         Width           =   1365
      End
      Begin VB.CommandButton CMDAccount8 
         Height          =   315
         Left            =   3225
         Picture         =   "PurchaseAccountSetup.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2910
         Width           =   315
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3555
         MaxLength       =   50
         TabIndex        =   30
         Top             =   2565
         Width           =   2865
      End
      Begin VB.TextBox txtAccount7 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   29
         Top             =   2580
         Width           =   1365
      End
      Begin VB.CommandButton CMDAccount7 
         Height          =   315
         Left            =   3210
         Picture         =   "PurchaseAccountSetup.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2565
         Width           =   315
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1440
         Width           =   2865
      End
      Begin VB.TextBox txtAccount4 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1470
         Width           =   1365
      End
      Begin VB.CommandButton Cmdaccount4 
         Height          =   315
         Left            =   3225
         Picture         =   "PurchaseAccountSetup.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
      End
      Begin VB.CommandButton CMDAccount6 
         Height          =   315
         Left            =   3225
         Picture         =   "PurchaseAccountSetup.frx":0BB6
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2175
         Width           =   315
      End
      Begin VB.TextBox txtAccount6 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2205
         Width           =   1365
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2175
         Width           =   2865
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1815
         Width           =   2865
      End
      Begin VB.TextBox txtaccount5 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1845
         Width           =   1365
      End
      Begin VB.CommandButton CmdAccount5 
         Height          =   315
         Left            =   3225
         Picture         =   "PurchaseAccountSetup.frx":0D28
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1815
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount3 
         Height          =   315
         Left            =   3225
         Picture         =   "PurchaseAccountSetup.frx":0E9A
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtaccount3 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1110
         Width           =   1365
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1080
         Width           =   2865
      End
      Begin VB.CommandButton CmdAccount2 
         Height          =   315
         Left            =   3225
         Picture         =   "PurchaseAccountSetup.frx":100C
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   315
      End
      Begin VB.TextBox txtaccount2 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   9
         Top             =   750
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   8
         Top             =   720
         Width           =   2865
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   7
         Top             =   345
         Width           =   2865
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Reset"
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   4875
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   5505
         TabIndex        =   5
         Top             =   4875
         Width           =   930
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Save"
         Height          =   330
         Left            =   4560
         TabIndex        =   4
         Top             =   4890
         Width           =   930
      End
      Begin VB.TextBox txtaccount1 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   0
         Top             =   375
         Width           =   1365
      End
      Begin VB.CommandButton CmdAccount1 
         Height          =   315
         Left            =   3225
         Picture         =   "PurchaseAccountSetup.frx":117E
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Purchase Act :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   48
         Top             =   4110
         Width           =   1740
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Misc Charge Act :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   43
         Top             =   3705
         Width           =   1740
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "W/H Tax Act :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   39
         Top             =   3345
         Width           =   1740
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Labour Charge Act :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   35
         Top             =   2970
         Width           =   1740
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost of Sale :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   31
         Top             =   2610
         Width           =   1260
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash Account :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   27
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Loading Charge Act :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   23
         Top             =   2235
         Width           =   1740
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Freight Account :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   19
         Top             =   1875
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sed Account :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   15
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "GST Account :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   11
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Flat Disc Account :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   3
         Top             =   405
         Width           =   1725
      End
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Misc Charge Act :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   30
      TabIndex        =   47
      Top             =   3990
      Width           =   1740
   End
End
Attribute VB_Name = "frmPOPurchaseAccountSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Dumy As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object
Public ls_transtype As String

Private Sub CMDAccount10_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount10
    Set PO_DESC = Text11
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount10) > 0 Then Call txtaccount10_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Cmdaccount4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount4
    Set PO_DESC = Text8
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount4) > 0 Then Call txtaccount4_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub CmdAccount5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount5
    Set PO_DESC = Text5
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount5) > 0 Then Call txtaccount5_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub CMDAccount6_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount6
    Set PO_DESC = Text3
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount6) > 0 Then Call txtaccount6_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub CMDAccount7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount7
    Set PO_DESC = Text7
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount7) > 0 Then Call txtaccount7_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub CMDAccount8_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount8
    Set PO_DESC = Text9
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount8) > 0 Then Call txtaccount8_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub CMDAccount9_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount9
    Set PO_DESC = Text10
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount9) > 0 Then Call txtaccount9_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command10_Click()
Call restvalues
End Sub

Private Sub restvalues()
Dim pr_dumyloadvalue As New Recordset
pr_dumyloadvalue.Open "Select * from PurchaseAccountSetup where compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyloadvalue.EOF Then
txtaccount1 = Trim(pr_dumyloadvalue("FlatDiscAccount") & "")
If txtaccount1 <> "" Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
txtaccount2 = Trim(pr_dumyloadvalue("GSTAccount") & "")
If txtaccount2 <> "" Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
txtaccount3 = Trim(pr_dumyloadvalue("SEdAccount") & "")
If txtaccount3 <> "" Then Call txtaccount3_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount4 = Trim(pr_dumyloadvalue("CashAccount") & "")
If txtAccount4 <> "" Then Call txtaccount4_KeyDown(vbKeyReturn, vbKeyShift)
txtaccount5 = Trim(pr_dumyloadvalue("OtherAccount1") & "")
If txtaccount5 <> "" Then Call txtaccount5_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount6 = Trim(pr_dumyloadvalue("OtherAccount2") & "")
If txtAccount6 <> "" Then Call txtaccount6_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount7 = Trim(pr_dumyloadvalue("CostofSale") & "")
If txtAccount7 <> "" Then Call txtaccount7_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount8 = Trim(pr_dumyloadvalue("OtherAccount3") & "")
If txtAccount8 <> "" Then Call txtaccount8_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount9 = Trim(pr_dumyloadvalue("WHTaxACT") & "")
If txtAccount9 <> "" Then Call txtaccount9_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount10 = Trim(pr_dumyloadvalue("MiscCharges") & "")
If txtAccount10 <> "" Then Call txtaccount10_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount11 = Trim(pr_dumyloadvalue("PurchaseAccount") & "")
If txtAccount11 <> "" Then Call txtaccount11_KeyDown(vbKeyReturn, vbKeyShift)

End If
pr_dumyloadvalue.Close
End Sub

Private Sub Command8_Click()
Dim ls_sql As String


ls_sql = "delete from PurchaseAccountSetup where compcode = '" & Gs_compcode & "' "
gc_dbcon.Execute ls_sql

ls_sql = "insert into  PurchaseAccountSetup (Compcode, FlatDiscAccount,GSTAccount,SEDAccount,CashAccount,OtherAccount1,OtherAccount2,CostofSale,OtherAccount3,WhTaxAct,MiscCharges,PurchaseAccount)"
ls_sql = ls_sql & " values ('" & Gs_compcode & "','" & txtaccount1 & "','" & txtaccount2 & "','" & txtaccount3 & "','" & txtAccount4 & "','" & txtaccount5 & "','" & txtAccount6 & "','" & txtAccount7 & "','" & txtAccount8 & "','" & txtAccount9 & "','" & txtAccount10 & "','" & txtAccount11 & "')"
gc_dbcon.Execute ls_sql

Call MsgBox("Successfully Updated !!!", vbInformation)
Call restvalues

End Sub

Private Sub Command9_Click()
txtaccount1 = ""
Text1.Text = ""

txtaccount2 = ""
Text2.Text = ""

txtaccount3 = ""
Text4.Text = ""

txtAccount4 = ""
Text8.Text = ""

txtaccount5 = ""
Text5.Text = ""

txtAccount6 = ""
Text3.Text = ""

txtAccount7 = ""
Text7.Text = ""

txtAccount8 = ""
Text9.Text = ""

txtAccount9 = ""
Text10.Text = ""

txtAccount10 = ""
Text11.Text = ""
End Sub

Private Sub Form_Load()
Call restvalues
End Sub
Private Sub CmdAccount1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount1
    Set PO_DESC = Text1
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount1) > 0 Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub
Private Sub txtaccount1_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount1 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtaccount1 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text1 = PR_Dumy("description")
            End If
         PR_Dumy.Close

End If

End Sub
Private Sub txtaccount9_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount9 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount9 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text10 = PR_Dumy("description")
                'txtAccount10.SetFocus
            End If
         PR_Dumy.Close

End If

End Sub
Private Sub txtaccount10_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount10 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount10 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text11 = PR_Dumy("description")
            End If
         PR_Dumy.Close
End If
End Sub
Private Sub CmdAccount2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount2
    Set PO_DESC = Text2
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount2) > 0 Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub
Private Sub txtaccount11_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount11 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount11 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text12 = PR_Dumy("description")
            End If
         PR_Dumy.Close
End If
End Sub
Private Sub CmdAccount11_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount11
    Set PO_DESC = Text12
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount11) > 0 Then Call txtaccount11_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub

Private Sub txtaccount2_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount2 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtaccount2 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text2 = PR_Dumy("description")
                
            End If
         PR_Dumy.Close

End If

End Sub
Private Sub CmdAccount3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount3
    Set PO_DESC = Text4
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount3) > 0 Then Call txtaccount3_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub

Private Sub txtaccount3_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount3 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtaccount3 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text4 = PR_Dumy("description")
            End If
         PR_Dumy.Close

End If

End Sub
Private Sub txtaccount4_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount4 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount4 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text8 = PR_Dumy("description")
                
            End If
         PR_Dumy.Close

End If

End Sub


Private Sub txtaccount5_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount5 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtaccount5 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text5 = PR_Dumy("description")
                
            End If
         PR_Dumy.Close

End If
End Sub
Private Sub txtaccount6_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount6 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount6 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text3 = PR_Dumy("description")
                
            End If
         PR_Dumy.Close

End If
End Sub
Private Sub txtaccount7_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount7 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount7 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text7 = PR_Dumy("description")
                
                
            End If
         PR_Dumy.Close

End If

End Sub
Private Sub txtaccount8_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount8 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount7 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text9 = PR_Dumy("description")
                
            End If
         PR_Dumy.Close

End If

End Sub

