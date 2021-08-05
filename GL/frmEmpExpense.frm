VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpExpense 
   Caption         =   "Employee Expenses"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpExpense.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8355
   StartUpPosition =   1  'CenterOwner
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
      Height          =   3090
      Left            =   45
      TabIndex        =   1
      Top             =   570
      Width           =   8265
      Begin VB.TextBox txtvoucherno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   6975
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   34
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   975
         Width           =   1185
      End
      Begin VB.TextBox txtvchrtype 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   975
         Width           =   660
      End
      Begin VB.CommandButton Command6 
         Height          =   315
         Left            =   2115
         Picture         =   "frmEmpExpense.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox txtexpensetype 
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
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   660
      End
      Begin VB.TextBox txtexpensedesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2475
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   5670
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Total Job Order Value"
         Top             =   930
         Width           =   1290
      End
      Begin VB.TextBox txtbankdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1665
         Width           =   3000
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
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1665
         Width           =   660
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   2115
         Picture         =   "frmEmpExpense.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1665
         Width           =   315
      End
      Begin VB.ComboBox txtpaymenttype 
         Height          =   330
         ItemData        =   "frmEmpExpense.frx":05EE
         Left            =   1440
         List            =   "frmEmpExpense.frx":05F8
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1305
         Width           =   2310
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
         Left            =   6585
         MaxLength       =   200
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1545
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1875
         Picture         =   "frmEmpExpense.frx":0608
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2595
         Width           =   315
      End
      Begin VB.TextBox txtacode 
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
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2595
         Width           =   435
      End
      Begin VB.TextBox txtaname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2595
         Width           =   1905
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   5715
         Picture         =   "frmEmpExpense.frx":077A
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2595
         Width           =   315
      End
      Begin VB.TextBox txtacode1 
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
         Left            =   5295
         MaxLength       =   3
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2610
         Width           =   435
      End
      Begin VB.TextBox txtaname1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6045
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2610
         Width           =   2115
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
         Height          =   315
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2550
         Picture         =   "frmEmpExpense.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   150
         Width           =   315
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   6660
         TabIndex        =   6
         Top             =   165
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23265281
         CurrentDate     =   37580
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   885
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   510
         Left            =   1425
         MaxLength       =   100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2040
         Width           =   6705
      End
      Begin Crystal.CrystalReport rptVoucher 
         Left            =   8370
         Top             =   180
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
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7470
         Top             =   -180
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Expense Code :"
         Height          =   255
         Left            =   60
         TabIndex        =   32
         Top             =   555
         Width           =   1350
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   28
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bank Code :"
         Height          =   210
         Index           =   2
         Left            =   525
         TabIndex        =   27
         Top             =   1695
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Payment Type :"
         Height          =   210
         Index           =   3
         Left            =   300
         TabIndex        =   26
         Top             =   1335
         Width           =   1110
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Instrument # :"
         Height          =   255
         Left            =   5505
         TabIndex        =   25
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Verified By :"
         Height          =   255
         Left            =   450
         TabIndex        =   18
         Top             =   2625
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Approved By :"
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   2625
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Expense #  :"
         Height          =   255
         Left            =   390
         TabIndex        =   10
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   630
         TabIndex        =   4
         Top             =   2055
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Expense Date :"
         Height          =   255
         Index           =   0
         Left            =   5475
         TabIndex        =   3
         ToolTipText     =   "Enter Value Date"
         Top             =   195
         Width           =   1140
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
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
               Picture         =   "frmEmpExpense.frx":0A5E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmpExpense.frx":0EB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmpExpense.frx":1306
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmpExpense.frx":175A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmpExpense.frx":1BAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmpExpense.frx":2002
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEmpExpense.frx":2756
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmEmpExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGRN As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Dumy As New Recordset
Dim PR_ICIssue As New Recordset
Dim PR_Branch As New Recordset
Dim ld_VluDate As Date
Dim ls_opt
Private Function maxtranscode() As String
PR_Dumy.Open "select max(transcode) as transcode from GL_EmpExpense where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & PR_Dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
PR_Dumy.Close
End Function
Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtacode
    Set PO_DESC = txtaname
    Gs_SQL = "Select ACode, Aname Description from PO_AuthorityPerson "
    Gs_FindFld = "Aname"
    Gs_OrderBy = "Order by AName"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Authority Person"
    MyLookupOLDB.Show 1
    
    If txtacode <> "" Then Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command6_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtexpensetype
    Set PO_DESC = txtexpensedesc
    Gs_SQL = "Select ECode,  Description from GL_ExenseType "
    Gs_FindFld = "Aname"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Expense Type"
    MyLookupOLDB.Show 1
    
    If txtexpensetype <> "" Then Call txtexpensetype_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub txtACode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtacode) <> "" And KeyCode = vbKeyReturn Then
        txtacode = DoPad(txtacode, 3)
        PR_Dumy.Open "Select * from PO_AuthorityPerson where Compcode  = '" & Gs_compcode & "' and Acode = '" & txtacode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Authority Code not found !!!", vbCritical)
            txtacode = ""
            txtaname = ""
            txtacode.SetFocus
        Else
            txtaname = PR_Dumy("aname")
            txtacode1.SetFocus
        End If
        PR_Dumy.Close

ElseIf Trim(txtacode) = "" And KeyCode = vbKeyReturn Then
        txtacode = ""
        txtaname = ""
End If

End Sub


Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtacode1
    Set PO_DESC = txtaname1
    Gs_SQL = "Select ACode, Aname Description from PO_AuthorityPerson "
    Gs_FindFld = "Aname"
    Gs_OrderBy = "Order by AName"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Authority Person"
    MyLookupOLDB.Show 1
    
    If txtacode1 <> "" Then Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)

End Sub
Private Sub txtACode1_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtacode1) <> "" And KeyCode = vbKeyReturn Then
        txtacode1 = DoPad(txtacode1, 3)
        PR_Dumy.Open "Select * from PO_AuthorityPerson where Compcode  = '" & Gs_compcode & "' and Acode = '" & txtacode1 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Authority Code not found !!!", vbCritical)
            txtacode1 = ""
            txtaname1 = ""
            txtacode1.SetFocus
        Else
            txtaname1 = PR_Dumy("aname")
        End If
        PR_Dumy.Close

ElseIf Trim(txtacode1) = "" And KeyCode = vbKeyReturn Then
        txtacode1 = ""
        txtaname1 = ""
End If
End Sub
Private Sub TxtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtpaymenttype.SetFocus
End Sub

Private Sub txtinstrumentNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then TxtRemarks.SetFocus
End Sub
Private Sub txtremarks_LostFocus()
If Trim(TxtRemarks) <> "" Then
    TxtRemarks = UCase(TxtRemarks)
End If
End Sub

Private Sub txtexpensetype_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtexpensetype) <> "" And KeyCode = vbKeyReturn Then
       txtexpensetype = DoPad(txtexpensetype, 3)
        PR_Dumy.Open "Select * from GL_ExenseType where Compcode  = '" & Gs_compcode & "' and Ecode = '" & txtexpensetype & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Expense Code not found !!!", vbCritical)
            txtexpensetype = ""
            txtexpensedesc = ""
            txtexpensetype.SetFocus
        Else
            txtexpensedesc = PR_Dumy("Description")
            If txtAmount.Enabled Then txtAmount.SetFocus
            TxtRemarks = UCase("Amount Paid to " & txtVendorDesc & " Against expense " & txtexpensedesc)
            
        End If
        PR_Dumy.Close
End If

End Sub


Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtTransNo
    Set PO_DESC = Text1
    Gs_SQL = "Select TransCode, Transdate from GL_EmpExpense "
    Gs_FindFld = "TransCode"
    Gs_OtherPara = "Where Compcode = '" & Gs_compcode & "'  and glstatus = 0"
    Gs_OrderBy = "Order by TransCode"

    MyLookupOLDB.Caption = "Vendor Advance"
    MyLookupOLDB.Show 1
    
    If txtTransNo <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_ICIssue, Me, txtTransNo, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
  ls_transtype = "D"
  
  
  SetToolBar(1) = True
  SetToolBar(2) = True
  SetToolBar(3) = True
  SetToolBar(4) = True
  
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  


  txtvaluedate.Value = Date
 
  

  
End Sub

Private Sub txtpaymenttype_Click()
If txtpaymenttype.Text = "Bank" Then
    txtbankcode.SetFocus
Else
   If TxtRemarks.Enabled Then TxtRemarks.SetFocus
End If
End Sub
Private Sub txtpaymenttype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtpaymenttype.Text = "Bank" Then
        txtbankcode.SetFocus
End If
End Sub
Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtacode.SetFocus

End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyReturn And Len(txtTransNo.Text) > 0 Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
         PR_ICIssue.Open "select * from GL_EmpExpense where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "'  ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       Select Case Mode
            Case "A"
                If Not PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   If txtTransNo.Enabled Then txtTransNo.SetFocus
                Else
                   txtvaluedate.SetFocus
                End If
            Case Else
                If PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txtTransNo.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtTransNo.SetFocus
                   End If
                End If
            End Select
  
     End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       txtTransNo.Enabled = False
       Command1.Enabled = False
       
       
    Else
      txtTransNo.Enabled = True
       txtTransNo.SetFocus
       Command1.Enabled = True
    End If
    If Button.Index = 7 Then
    
    End If
    
    If PB_BlnkGRN And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_ICIssue, Me, txtTransNo, txtTransNo, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
       txtTransNo = maxtranscode
      txtexpensetype.SetFocus
    End If
End Sub


Public Sub SaveValues()
'On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String

gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
              gc_dbcon.Execute "DELETE FROM GL_EmpExpense WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtTransNo) & "'"
           Case Else
                If Mode = "E" Then
                    gc_dbcon.Execute "DELETE FROM GL_EmpExpense WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtTransNo) & "' "
                End If

                If Mode = "A" Then
                    txtTransNo = maxtranscode
                End If
                    
                      ls_sql = "INSERT into GL_EmpExpense(Compcode,branchcode, TransCode, TransDate, ExpCode, Remarks,vcode,acode,Amount,PaymentType,Bankcode,Instrno)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtexpensetype & "','" & RepApp(TxtRemarks) & "','" & txtacode & "','" & txtacode1 & "'," & Val(txtAmount) & "," & Val(txtpaymenttype.ListIndex) & ",'" & Trim(txtbankcode) & "','" & Trim(txtinstrumentNo) & "')"
                      gc_dbcon.Execute ls_sql
                
     End Select
gc_dbcon.CommitTrans
            If Mode = "A" Or Mode = "E" Then
            
            ls_sql = "DELETE FROM Gl_Trans WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' AND Voucher_No = '" & txtvoucherno & "' AND VchrType = '" & txtvchrtype & "' and month(value_date) = " & Month(txtvaluedate.Value) & " and year(value_date) = " & Year(txtvaluedate.Value) & ""
            gc_dbcon.Execute ls_sql
              
            ls_sql = "DELETE FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' AND Voucher_No = '" & txtvoucherno & "' AND VchrType = '" & txtvchrtype & "' and month(value_date) = " & Month(txtvaluedate.Value) & " and year(value_date) = " & Year(txtvaluedate.Value) & ""
            gc_dbcon.Execute ls_sql
              ld_VluDate = txtvaluedate
              Call PostExpenseVoucher(Gs_compcode, txtTransNo)
              ls_opt = SetErr("Print Voucher ?.", vbYesNo)
              If ls_opt = vbYes Then Call setprint
            End If

'If Mode <> "D" Then
'   ls_opt = MsgBox("Print Job Order Note ?.", vbYesNo)
'   If ls_opt = vbYes Then Call PrintIssuenote
'End If

If Mode = "A" Then
    txtTransNo = maxtranscode
End If

Exit Sub
RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Public Sub ClearVal()
End Sub

Private Sub PrintIssuenote()
On Error GoTo LocalErr

   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "ContJobOrderNote.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Job Order'"
        .SelectionFormula = "{PO_DemandNoteDetail.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.transcode} = '" & Trim(txtTransNo) & "' "
        .Action = 1
   End With
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub

Private Sub SetVal()
     txtvaluedate = PR_ICIssue("Transdate")
     txtexpensetype = Trim(PR_ICIssue("ExpCode") & "")
     Call txtexpensetype_KeyDown(vbKeyReturn, vbKeyShift)
     TxtRemarks = Trim(PR_ICIssue("Remarks") & "")
     txtacode = Trim(PR_ICIssue("VCode") & "")
     Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
     txtacode1 = Trim(PR_ICIssue("ACode") & "")
     Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)
     txtAmount = Val(0 & PR_ICIssue("Amount"))
     txtpaymenttype.ListIndex = Val(0 & PR_ICIssue("PaymentType"))
     txtbankcode = Trim(PR_ICIssue("Bankcode") & "")
     If txtbankcode <> "" Then Call txtbankcode_KeyDown(vbKeyReturn, vbKeyShift)
     txtinstrumentNo = Trim(PR_ICIssue("instrno") & "")
     txtvchrtype = Trim(PR_ICIssue("Vchrtype") & "")
     txtvoucherno = Trim(PR_ICIssue("VoucherNo") & "")
End Sub
Private Sub setprint()
On Error GoTo LocalErr
Dim ls_BranchName As String
Dim ls_VchDesc As String
Dim PR_Dumy As New Recordset
txtvchrtype = Gs_Vchrtype
txtvoucherno = Gs_VoucherNo
PR_Dumy.Open "Select * from sysbranch where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
ls_BranchName = PR_Dumy("BranchDesc")
End If
PR_Dumy.Close
PR_Dumy.Open "Select * from GlVchrType where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' and vchrtype = '" & txtvchrtype & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
ls_VchDesc = PR_Dumy("VchrDescrip")
End If
PR_Dumy.Close
   
   
   With CrystalReport1
        .ReportFileName = App.Path & Gs_GlRepoPath & "\Vchr_Print.RPT"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & ls_VchDesc & "'"
        .Formulas(5) = "BranchName = '" & Gs_BranchCode + "-" + ls_BranchName & "'"
        .SelectionFormula = "{Gl_Trans.Voucher_No} = '" & Trim(txtvoucherno) & "' and {Gl_Trans.BranchCode} = '" & Gs_BranchCode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.VchrType} = '" & Trim(txtvchrtype) & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.CompCode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.Value_Date} = Date(" & Year(ld_VluDate) & "," & Month(ld_VluDate) & "," & Day(ld_VluDate) & ")"
        .Formulas(2) = "Sig1 = '" & Gc_UserName & "'"
        .Formulas(3) = "Sig2 = '" & Gs_Sign2 & "'"
        .Formulas(4) = "Sig3 = '" & Gs_Sign3 & "'"
        .Action = 1
   End With
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)

End Sub


Private Sub Command4_Click()
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

Private Sub txtbankcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtbankcode <> "" Then
    txtbankcode = DoPad(txtbankcode, txtbankcode.MaxLength)
    
        ls_sql = "Select bankcode,bankname as Description from SysBanks where bankcode = '" & txtbankcode & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Bank code not found", vbCritical)
                Cancel = True
            Else
                txtbankdesc = PR_Dumy("description")
                If txtinstrumentNo.Enabled Then txtinstrumentNo.SetFocus
            End If
         PR_Dumy.Close
    
End If

End Sub

Public Function ChkInputs() As Boolean
 Dim lb_opt As Boolean
    If Trim(txtexpensetype) = "" Then
      Call MsgBox("Enter/Select Expense Code !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(txtAmount) = "" Then
      Call MsgBox("Enter Advance Amount !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(TxtRemarks) = "" Then
      Call MsgBox("Enter Remarks !!!", vbCritical)
      ChkInputs = False
      TxtRemarks.SetFocus
    ElseIf Trim(txtacode) = "" Then
      Call MsgBox("Enter/Select Verified Code !!!", vbCritical)
      ChkInputs = False
      txtacode.SetFocus
    ElseIf Trim(txtacode1) = "" Then
      Call MsgBox("Enter/Select Approved Code !!!", vbCritical)
      ChkInputs = False
      txtacode1.SetFocus
    Else
      ChkInputs = True
    End If
End Function

Public Sub FrmRefresh()
End Sub
Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtVendorCode.SetFocus
    End If
End Sub

