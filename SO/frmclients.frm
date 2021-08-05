VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmclients 
   Caption         =   "Clients"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmclients.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7290
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
      Height          =   4860
      Left            =   15
      TabIndex        =   10
      Top             =   600
      Width           =   7245
      Begin VB.CheckBox chkapproved 
         Alignment       =   1  'Right Justify
         Caption         =   "Approved :"
         Height          =   330
         Left            =   6000
         TabIndex        =   45
         Top             =   4440
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txtopeningbal 
         Height          =   315
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   4425
         Width           =   1725
      End
      Begin VB.CommandButton Command5 
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
         Left            =   1710
         Picture         =   "frmclients.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   41
         Tag             =   "SKIP"
         Top             =   2160
         Width           =   315
      End
      Begin VB.TextBox txtsector 
         Height          =   330
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   40
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtsectordesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2055
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   39
         Tag             =   "SKIP"
         Top             =   2160
         Width           =   5100
      End
      Begin VB.TextBox txtstreg 
         Height          =   315
         Left            =   1215
         MaxLength       =   255
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4050
         Width           =   5940
      End
      Begin VB.TextBox txtcountryName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   36
         Tag             =   "SKIP"
         Top             =   2550
         Width           =   5100
      End
      Begin VB.TextBox txtsub0 
         Height          =   315
         Left            =   1215
         MaxLength       =   13
         TabIndex        =   35
         Top             =   3675
         Width           =   1515
      End
      Begin VB.TextBox txtSubDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3075
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3660
         Width           =   4110
      End
      Begin VB.CommandButton Command4 
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
         Left            =   2745
         Picture         =   "frmclients.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   3660
         Width           =   315
      End
      Begin VB.TextBox txtcountry 
         Height          =   330
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   30
         Top             =   2550
         Width           =   495
      End
      Begin VB.CommandButton Command3 
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
         Left            =   1725
         Picture         =   "frmclients.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "SKIP"
         Top             =   2550
         Width           =   315
      End
      Begin VB.TextBox txttehseelname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   26
         Tag             =   "SKIP"
         Top             =   3285
         Width           =   5115
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1740
         Picture         =   "frmclients.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "SKIP"
         Top             =   3300
         Width           =   315
      End
      Begin VB.TextBox txttehseel 
         Height          =   330
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   24
         Top             =   3300
         Width           =   495
      End
      Begin VB.TextBox txtcityname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         Tag             =   "SKIP"
         Top             =   2925
         Width           =   5115
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
         Left            =   1725
         Picture         =   "frmclients.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "SKIP"
         Top             =   2925
         Width           =   315
      End
      Begin VB.TextBox txtCity 
         Height          =   330
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   21
         Top             =   2925
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtpagrdate 
         Height          =   315
         Left            =   4755
         TabIndex        =   7
         Top             =   1770
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   37848
      End
      Begin VB.ComboBox txttype 
         Height          =   330
         ItemData        =   "frmclients.frx":0A44
         Left            =   1215
         List            =   "frmclients.frx":0A4B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1770
         Width           =   2415
      End
      Begin VB.TextBox txtmobile 
         Height          =   315
         Left            =   1215
         MaxLength       =   25
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2415
      End
      Begin VB.TextBox txtphoneres 
         Height          =   315
         Left            =   4740
         MaxLength       =   25
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2415
      End
      Begin VB.TextBox txtphoneoffice 
         Height          =   315
         Left            =   4725
         MaxLength       =   25
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1020
         Width           =   2430
      End
      Begin VB.TextBox txtnic 
         Height          =   315
         Left            =   1215
         MaxLength       =   25
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1005
         Width           =   2400
      End
      Begin VB.TextBox txtaddress 
         Height          =   315
         Left            =   1215
         MaxLength       =   200
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   615
         Width           =   5940
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   2910
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   4245
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
         Left            =   1995
         Picture         =   "frmclients.frx":0A59
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtClientCode 
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   6
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
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Opening Bal :"
         Height          =   210
         Left            =   225
         TabIndex        =   44
         Top             =   4455
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sector Code :"
         Height          =   210
         Left            =   180
         TabIndex        =   42
         Top             =   2205
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "St.Reg# :"
         Height          =   195
         Left            =   210
         TabIndex        =   37
         Top             =   4095
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Gl Account# :"
         Height          =   210
         Left            =   210
         TabIndex        =   34
         Top             =   3705
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Country :"
         Height          =   210
         Left            =   525
         TabIndex        =   31
         Top             =   2595
         Width           =   660
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tehseel :"
         Height          =   210
         Left            =   540
         TabIndex        =   28
         Top             =   3345
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "City :"
         Height          =   210
         Left            =   795
         TabIndex        =   27
         Top             =   2970
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type :"
         Height          =   210
         Index           =   8
         Left            =   720
         TabIndex        =   20
         Top             =   1815
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agr.Date :"
         Height          =   210
         Index           =   7
         Left            =   3975
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mobile :"
         Height          =   210
         Index           =   6
         Left            =   630
         TabIndex        =   18
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Phone Res :"
         Height          =   210
         Index           =   5
         Left            =   3840
         TabIndex        =   17
         Top             =   1410
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Phone Office :"
         Height          =   210
         Index           =   4
         Left            =   3675
         TabIndex        =   16
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NTN# :"
         Height          =   210
         Index           =   3
         Left            =   750
         TabIndex        =   15
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         Height          =   210
         Index           =   2
         Left            =   450
         TabIndex        =   14
         Top             =   645
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   210
         Index           =   0
         Left            =   2370
         TabIndex        =   13
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   705
         TabIndex        =   11
         Top             =   270
         Width           =   465
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
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
               Picture         =   "frmclients.frx":0BCB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":101F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":1473
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":18C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":1D1B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":216F
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":28C3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmclients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkSupp As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Supplier As Recordset
Dim PR_Cities As New Recordset
Dim PR_Country As New Recordset
Dim PR_GlSub2 As New Recordset
Dim PR_Tehseels As New Recordset
Dim PR_Sector As New Recordset
Dim PR_Dumy As New Recordset
Dim pr_dumy1 As New Recordset
Dim PR_DumyBSetup As New Recordset
Dim ls_bnsicode As String
Dim ls_bcode As String
Dim ls_bncode As String


Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtClientCode
    Set PO_DESC = txtDesc
    GoTop PR_Supplier
    MyLookup.Caption = "Suppliers"
    MyLookup.FillGrid PR_Supplier, "ClientCode", "Description", 5
    MyLookup.Show 1
    
    If Len(txtClientCode) > 0 Then txtClientCode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command1_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCity
    Set PO_DESC = txtcityname
    
    GoTop PR_Cities
    MyLookup.Caption = "Cities"
    MyLookup.FillGrid PR_Cities, "CityCode", "CityName", txtCity.MaxLength
    MyLookup.Show 1
    
    If Len(txtCity) > 0 Then txtCity_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttehseel
    Set PO_DESC = txttehseelname
    
    GoTop PR_Tehseels
    PR_Tehseels.Filter = "CityCode = '" & txtCity.Text & "'"
    MyLookup.Caption = "Tehseels"
    MyLookup.FillGrid PR_Tehseels, "tehseelCode", "tehseelName", txttehseel.MaxLength
    MyLookup.Show 1
    PR_Tehseels.Filter = adFilterNone
    If Len(txttehseel) > 0 Then txttehseel_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcountry
    Set PO_DESC = txtcountryName
    
    GoTop PR_Country
    MyLookup.Caption = "Country"
    MyLookup.FillGrid PR_Country, "CountryCode", "CountryName", txtcountry.MaxLength
    MyLookup.Show 1
    
    If Len(txtcountry) > 0 Then txtcountry_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command4_Click()
Dim ln_SetLen As Integer
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsub0
    Set PO_DESC = txtSubDesc
    
    GoTop PR_GlSub2
    
    Gs_SQL = "Select accountno  'Account No', Acct_Desc  'Description' from  gl_detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc"
    MyLookupOLDB.Caption = "Sub Accounts."
    MyLookupOLDB.Show 1
   If Len(txtsub0) > 0 Then txtsub0_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsector
    Set PO_DESC = txtsectordesc
    
    GoTop PR_Sector
    MyLookup.Caption = "Sectors"
    MyLookup.FillGrid PR_Sector, "sectorCode", "sectordesc", txtsector.MaxLength
    MyLookup.Show 1
    
    If Len(txtsector) > 0 Then txtsector_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub txtDesc_Change()
If txtDesc <> "" Then
txtSubDesc = txtDesc
End If
End Sub

Private Sub txtsector_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtsector.Text = DoPad(txtsector.Text, txtsector.MaxLength)
      If MySeek(txtsector.Text, "sectorCode", PR_Sector) Then
           txtsectordesc = PR_Sector("sectordesc")
          If txtcountry.Enabled Then txtcountry.SetFocus
        Else
            Call SetErr(Gs_RecNFMsg, vbCritical)
            If txtsector.Enabled Then txtsector.SetFocus
      End If
ElseIf KeyCode = vbKeyPageUp Then
       dtpagrdate.SetFocus
ElseIf KeyCode = vbKeyF12 Then
        Call Command5_Click
End If
End Sub

Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtCity.Text = DoPad(txtCity.Text, txtCity.MaxLength)
      If MySeek(txtCity.Text, "CityCode", PR_Cities) Then
           txtcityname = PR_Cities("Cityname")
           If txttehseel.Enabled Then txttehseel.SetFocus
        Else
            Call SetErr(Gs_RecNFMsg, vbCritical)
            If txtCity.Enabled Then txtCity.SetFocus
      End If
ElseIf KeyCode = vbKeyPageUp Then
       txtcountry.SetFocus
ElseIf KeyCode = vbKeyF12 Then
        Call Command1_Click
End If
End Sub

Private Sub txtcountry_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtcountry.Text = DoPad(txtcountry.Text, txtcountry.MaxLength)
      If MySeek(txtcountry.Text, "CountryCode", PR_Country) Then
           txtcountryName = PR_Country("Countryname")
           If txtCity.Enabled Then txtCity.SetFocus
        Else
            Call SetErr(Gs_RecNFMsg, vbCritical)
            If txtcountry.Enabled Then txtcountry.SetFocus
      End If
ElseIf KeyCode = vbKeyPageUp Then
       dtpagrdate.SetFocus
ElseIf KeyCode = vbKeyF12 Then
        Call Command3_Click
End If
End Sub



Private Sub txtsub0_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Val(txtsub0.Text) <> 0 Then
         
         lb_found = MySeek(txtsub0.Text, "accountno", PR_GlSub2)
        
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtsub0.SetFocus
             txtSubDesc.Text = txtDesc
         Else
             txtSubDesc.Text = PR_GlSub2("Acct_Desc")
             If txtstreg.Enabled Then txtstreg.SetFocus
         End If
 ElseIf KeyCode = vbKeyPageUp Then
    txttehseel.SetFocus
 ElseIf KeyCode = vbKeyF12 Then
    Call Command4_Click
 End If
End Sub

Private Sub txttehseel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
      txttehseel.Text = DoPad(txttehseel.Text, txttehseel.MaxLength)
      
      If MySeek(Trim(txtCity) + Trim(txttehseel.Text), "findfld", PR_Tehseels) Then
           txttehseelname = PR_Tehseels("TehseelName")
           If txtsub0.Enabled Then
            txtsub0.SetFocus
           Else
            txtstreg.SetFocus
           End If
      Else
           Call SetErr(Gs_RecNFMsg, vbCritical)
           If txttehseel.Enabled Then txttehseel.SetFocus
      End If
ElseIf KeyCode = vbKeyPageUp Then
    txtCity.SetFocus
ElseIf KeyCode = vbKeyF12 Then
    Call Command2_Click
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Supplier, frmSupplier, txtClientCode, txtDesc, Para_Rs, "IC_PartyCnt", 6, "ClientCode", "Description", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SRCST00001")
  SetToolBar(2) = chkRights("SRCST00002")
  SetToolBar(3) = chkRights("SRCST00003")
  SetToolBar(4) = chkRights("SRCST00004")
  
  'Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  'Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  'Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  'Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  txttype = "Customer"
  
  Set PR_Supplier = New Recordset
   
  PR_Supplier.Open "Select * from IC_Clients where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Cities.Open "Select * from cities order by CityCode ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Country.Open "Select * from Country order by CountryCode ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Sector.Open "Select * from Ic_sectors order by SectorCode ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Tehseels.Open "Select CityCode+TehseelCode as findfld,*  from tehseels order by citycode,tehseelcode ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PB_BlnkSupp = IIf(PR_Supplier.EOF, True, False)
  PR_GlSub2.Open "Select * from Gl_detail where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Supplier.Close
    PR_Cities.Close
    PR_Country.Close
    PR_Tehseels.Close
    PR_GlSub2.Close
    PR_Sector.Close
End Sub


Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtaddress.SetFocus
    If KeyCode = vbKeyPageUp And txtClientCode.Enabled Then txtClientCode.SetFocus
End Sub
Private Sub txtaddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtnic.SetFocus
    If KeyCode = vbKeyPageUp Then txtDesc.SetFocus
End Sub
Private Sub txtNIC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtphoneoffice.SetFocus
    If KeyCode = vbKeyPageUp Then txtaddress.SetFocus
End Sub
Private Sub txtphoneoffice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtmobile.SetFocus
    If KeyCode = vbKeyPageUp Then txtnic.SetFocus
End Sub
Private Sub txtphoneres_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txttype.SetFocus
    If KeyCode = vbKeyPageUp Then txtmobile.SetFocus
End Sub
Private Sub txtMobile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtphoneres.SetFocus
    If KeyCode = vbKeyPageUp Then txtphoneoffice.SetFocus
End Sub
Private Sub dtpagrdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtsector.SetFocus
    If KeyCode = vbKeyPageUp Then txttype.SetFocus
End Sub
Private Sub txttype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpagrdate.SetFocus
    If KeyCode = vbKeyPageUp Then txtphoneres.SetFocus
End Sub

Private Sub txtClientCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtClientCode.Text) > 0 And Val(txtClientCode.Text) > 0 Then
          PR_Supplier.Requery
          
         txtClientCode.Text = DoPad(UCase(txtClientCode.Text), txtClientCode.MaxLength)
         lb_found = MySeek(txtClientCode.Text, "ClientCode", PR_Supplier)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   If txtClientCode.Enabled Then txtClientCode.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                   Call ClearVal
                   If txtClientCode.Enabled Then txtClientCode.SetFocus
                Else
                   Call SetVal
                End If
            End Select
            
       End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
    Else
      cmdLookup.Enabled = True
    End If
    
    If PB_BlnkSupp And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Supplier, Me, txtClientCode, txtDesc, Para_Rs, "IC_PartyCnt", 6, "ClientCode", "Description", 1, False, Toolbar1)
    End If
    txtsub0.Enabled = True
    If Mode = "A" Then
       txtClientCode = maxtranscode
       txtsub0 = Clientcode + Right(Trim(txtClientCode), 4)
       txtsub0.Enabled = False
       txtDesc.SetFocus
    End If

    
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
Dim ln_cnt As Integer
Dim ls_CodeID As String
PB_BlnkSupp = False


gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT into IC_Clients(Compcode,ClientCode,Description,Address,NicNo,Phoneoffice,Phoneres,Mobile,AgrDate,CodeID,CountryCode,CityCode,TehseelCode,GlAccountNo,stregno,sectorcode,openingbalance) VALUES ('" & Gs_compcode & "','" & txtClientCode.Text & "','" & txtDesc.Text & "','" & txtaddress.Text & "','" & txtnic.Text & "','" & txtphoneoffice.Text & "','" & txtphoneres.Text & "','" & txtmobile.Text & "','" & Format(dtpagrdate, "YYYY/MM/DD") & "','" & ls_CodeID & "','" & Trim(txtcountry) & "','" & Trim(txtCity) & "','" & Trim(txttehseel) & "','" & Trim(txtsub0) & "','" & txtstreg & "','" & txtsector & "'," & Val(txtopeningbal) & ")"
              cmdLookup.Enabled = False
              
              If chkapproved.Value = 1 Then
                gc_dbcon.Execute "INSERT into IC_Clients(Compcode,ClientCode,Description,Address,NicNo,Phoneoffice,Phoneres,Mobile,AgrDate,CodeID,CountryCode,CityCode,TehseelCode,GlAccountNo,stregno,sectorcode,openingbalance) VALUES ('" & Gs_scompcode & "','" & txtClientCode.Text & "','" & txtDesc.Text & "','" & txtaddress.Text & "','" & txtnic.Text & "','" & txtphoneoffice.Text & "','" & txtphoneres.Text & "','" & txtmobile.Text & "','" & Format(dtpagrdate, "YYYY/MM/DD") & "','" & ls_CodeID & "','" & Trim(txtcountry) & "','" & Trim(txtCity) & "','" & Trim(txttehseel) & "','" & Trim(txtsub0) & "','" & txtstreg & "','" & txtsector & "'," & Val(txtopeningbal) & ")"
              End If
              
           Case "E"
              gc_dbcon.Execute "UPDATE IC_Clients SET Description= '" & txtDesc.Text & "',Address= '" & txtaddress.Text & "',NicNo= '" & txtnic.Text & "',Phoneoffice= '" & txtphoneoffice.Text & "',PhoneRes= '" & txtphoneres.Text & "',Mobile= '" & txtmobile.Text & "',AgrDate= '" & Format(dtpagrdate, "YYYY/MM/DD") & "', CodeID ='" & ls_CodeID & "', countryCode  = '" & txtcountry & "', CityCode  = '" & txtCity & "', TehseelCode  = '" & txttehseel & "', GlAccountNo  = '" & Trim(txtsub0) & "',stregno = '" & txtstreg & "',sectorcode = '" & txtsector & "',openingbalance = " & Val(txtopeningbal) & " WHERE  ClientCode= '" & txtClientCode.Text & "' and Compcode = '" & Gs_compcode & "'"
              
           Case "D"
             gc_dbcon.Execute "DELETE FROM IC_Clients WHERE ClientCode = '" & txtClientCode.Text & "' and Compcode = '" & Gs_compcode & "'"
           
     End Select

'glCode

PR_Dumy.Open "Select * from Gl_Detail where compcode = '" & Gs_compcode & "' and accountno = '" & txtsub0 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If PR_Dumy.EOF Then

ls_sql = "INSERT into Gl_detail(compcode,Acct_sub,Acct_Detail,AccountNo,Acct_desc,crncy_code,Acct_Base,Acct_Type,Acct_Status,userid,adddate,addtime) VALUES ('" & Gs_compcode & "','" & Left(txtsub0, 9) & "', '" & Right(txtClientCode, 4) & " ','" & txtsub0 & "','" & txtDesc & "','PKR','B','G','D','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
gc_dbcon.Execute ls_sql

End If
PR_Dumy.Close

'end gl Code

'Balance Sheet





PR_Dumy.Open "Select * from GL_Bsheet3Detail where compcode = '" & Gs_compcode & "' and accountno = '" & txtsub0 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If PR_Dumy.EOF Then

    
    

    PR_DumyBSetup.Open "Select * from ClientBSSetup where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    If Not PR_DumyBSetup.EOF Then
    ls_bcode = Trim(PR_DumyBSetup("BSMainHead") & "")
    ls_bncode = Trim(PR_DumyBSetup("BSNotes") & "")
    End If
    PR_DumyBSetup.Close

    ls_bnsicode = maxBSCode(Gs_compcode)
    
    If ls_bcode <> "" Then
        ls_sql = "insert into  gl_bsheet3 (compcode,bcode,bncode,bnicode,bnidesc) values ('" & Gs_compcode & "','" & ls_bcode & "','" & ls_bncode & "','" & ls_bnsicode & "','" & txtDesc & "')"
        gc_dbcon.Execute ls_sql
    
        ls_sql = "insert into  gl_bsheet3detail (compcode,bcode,bncode,bnicode,accountno) values ('" & Gs_compcode & "','" & ls_bcode & "','" & ls_bncode & "','" & ls_bnsicode & "','" & txtsub0 & "')"
        gc_dbcon.Execute ls_sql
    Else
       Call MsgBox("Client Balance Sheet not setup", vbCritical)
    End If
    
End If
PR_Dumy.Close




If chkapproved.Value = 1 Then
'GL CODE
PR_Dumy.Open "Select * from Gl_Detail where compcode = '" & Gs_scompcode & "' and accountno = '" & txtsub0 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If PR_Dumy.EOF Then

ls_sql = "INSERT into Gl_detail(compcode,Acct_sub,Acct_Detail,AccountNo,Acct_desc,crncy_code,Acct_Base,Acct_Type,Acct_Status,userid,adddate,addtime) VALUES ('" & Gs_scompcode & "','" & Left(txtsub0, 9) & "', '" & Right(txtClientCode, 4) & " ','" & txtsub0 & "','" & txtDesc & "','PKR','B','G','D','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
gc_dbcon.Execute ls_sql

End If
PR_Dumy.Close
'END GL CODE

'Balance Sheet

PR_Dumy.Open "Select * from GL_Bsheet3Detail where compcode = '" & Gs_scompcode & "' and accountno = '" & txtsub0 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If PR_Dumy.EOF Then

    PR_DumyBSetup.Open "Select * from ClientBSSetup where compcode = '" & Gs_scompcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    If Not PR_DumyBSetup.EOF Then
    ls_bcode = Trim(PR_DumyBSetup("BSMainHead") & "")
    ls_bncode = Trim(PR_DumyBSetup("BSNotes") & "")
    End If
    PR_DumyBSetup.Close

    ls_bnsicode = maxBSCode(Gs_scompcode)
    
    If ls_bcode <> "" Then
        ls_sql = "insert into  gl_bsheet3 (compcode,bcode,bncode,bnicode,bnidesc) values ('" & Gs_scompcode & "','" & ls_bcode & "','" & ls_bncode & "','" & ls_bnsicode & "','" & txtDesc & "')"
        gc_dbcon.Execute ls_sql
    
        ls_sql = "insert into  gl_bsheet3detail (compcode,bcode,bncode,bnicode,accountno) values ('" & Gs_scompcode & "','" & ls_bcode & "','" & ls_bncode & "','" & ls_bnsicode & "','" & txtsub0 & "')"
        gc_dbcon.Execute ls_sql
    Else
       Call MsgBox("Client Balance Sheet not setup", vbCritical)
    End If
    
End If
PR_Dumy.Close


End If

gc_dbcon.CommitTrans
PR_Supplier.Requery
PR_GlSub2.Requery




If Mode = "A" Then
       txtClientCode = maxtranscode
       txtDesc.SetFocus
End If


Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Function maxBSCode(cmpCode As String) As String


PR_DumyBSetup.Open "Select * from ClientBSSetup where compcode = '" & cmpCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    If Not PR_DumyBSetup.EOF Then
    ls_bcode = Trim(PR_DumyBSetup("BSMainHead") & "")
    ls_bncode = Trim(PR_DumyBSetup("BSNotes") & "")
    End If
PR_DumyBSetup.Close


pr_dumy1.Open "select max(bnicode) as bncode from Gl_BSheet3 where BCODE = '" & ls_bcode & "' and bncode = '" & ls_bncode & "' and compcode = '" & cmpCode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
    maxBSCode = DoPad(Trim(str(Val(0 & pr_dumy1("bncode")) + 1)), 4)
Else
    maxBSCode = DoPad(Trim(str(1)), txtbnscode.MaxLength)
End If
pr_dumy1.Close

End Function

Public Sub ClearVal()
     txtClientCode = ""
     txtDesc = ""
     txtCrValue = ""
     txtCrQty = ""
End Sub
Private Sub txtstreg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtopeningbal.SetFocus
End Sub

Private Sub SetVal()
     Dim ln_cnt As Integer
     txtDesc = PR_Supplier("Description")
     txtaddress = Trim(PR_Supplier("Address") & "")
     txtnic = Trim(PR_Supplier("NicNo") & "")
     txtphoneoffice = Trim(PR_Supplier("Phoneoffice") & "")
     txtphoneres = Trim(PR_Supplier("PhoneRes") & "")
     txtmobile = Trim(PR_Supplier("Mobile") & "")
     dtpagrdate = PR_Supplier("agrdate")
     txtcountry = PR_Supplier("CountryCode") & ""
     txtsector = PR_Supplier("sectorcode") & ""
     txtstreg = PR_Supplier("stregno") & ""
     txtopeningbal = Val(PR_Supplier("openingbalance"))
     If txtcountry <> "" Then Call txtcountry_KeyDown(vbKeyReturn, vbKeyShift)
     txtCity = PR_Supplier("CityCode") & ""
     If txtCity <> "" Then Call txtCity_KeyDown(vbKeyReturn, vbKeyShift)
     txttehseel = PR_Supplier("TehseelCode") & ""
     If txttehseel <> "" Then Call txttehseel_KeyDown(vbKeyReturn, vbKeyShift)
     txtsub0 = Trim(PR_Supplier("GlAccountNo") & "")
     If txtsub0 <> "" Then Call txtsub0_KeyDown(vbKeyReturn, vbKeyShift)
     If txtsector <> "" Then Call txtsector_KeyDown(vbKeyReturn, vbKeyShift)
     
End Sub
Private Function maxtranscode() As String
PR_Dumy.Open "select max(ClientCode) as transcode from ic_clients where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & PR_Dumy("transcode")) + 1)), 6)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 6)
End If
PR_Dumy.Close
End Function


Public Function ChkInputs() As Boolean
    If Len(txtClientCode.Text) = txtClientCode.MaxLength And txtDesc.Text <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtDesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub
