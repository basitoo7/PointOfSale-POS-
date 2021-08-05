VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClients 
   Caption         =   "Clients"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   7530
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDiscPer 
      Height          =   315
      Left            =   3360
      MaxLength       =   255
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   5880
      Width           =   705
   End
   Begin VB.TextBox txtBillCopy 
      Height          =   315
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5880
      Width           =   705
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
      Height          =   5820
      Left            =   15
      TabIndex        =   8
      Top             =   585
      Width           =   7485
      Begin VB.CheckBox ChkDiscAllow 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Discount Not Allow"
         Height          =   375
         Left            =   5040
         TabIndex        =   51
         Top             =   5280
         Width           =   2295
      End
      Begin VB.TextBox txtdeliveryAddress 
         Height          =   360
         Left            =   1425
         MaxLength       =   255
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   975
         Width           =   5940
      End
      Begin VB.TextBox txtwebsite 
         Height          =   315
         Left            =   5010
         MaxLength       =   255
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   4815
         Width           =   2355
      End
      Begin VB.TextBox txtContactperson 
         Height          =   315
         Left            =   1425
         MaxLength       =   25
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2145
         Width           =   2415
      End
      Begin VB.TextBox txtClientCode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1425
         MaxLength       =   6
         TabIndex        =   43
         Tag             =   "SKIP"
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtemail 
         Height          =   315
         Left            =   1425
         MaxLength       =   255
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   4800
         Width           =   2745
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
         Left            =   1920
         Picture         =   "frmclients.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "SKIP"
         Top             =   2535
         Width           =   315
      End
      Begin VB.TextBox txtsector 
         Height          =   330
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   38
         Top             =   2535
         Width           =   495
      End
      Begin VB.TextBox txtsectordesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2265
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   37
         Tag             =   "SKIP"
         Top             =   2535
         Width           =   5100
      End
      Begin VB.TextBox txtstreg 
         Height          =   315
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4425
         Width           =   5940
      End
      Begin VB.TextBox txtcountryName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         Tag             =   "SKIP"
         Top             =   2925
         Width           =   5100
      End
      Begin VB.TextBox txtsub0 
         Height          =   315
         Left            =   1425
         MaxLength       =   13
         TabIndex        =   33
         Top             =   4050
         Width           =   1515
      End
      Begin VB.TextBox txtSubDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3285
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4035
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
         Left            =   2955
         Picture         =   "frmclients.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   4035
         Width           =   315
      End
      Begin VB.TextBox txtcountry 
         Height          =   330
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   28
         Top             =   2925
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
         Left            =   1935
         Picture         =   "frmclients.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "SKIP"
         Top             =   2925
         Width           =   315
      End
      Begin VB.TextBox txttehseelname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   24
         Tag             =   "SKIP"
         Top             =   3660
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
         Left            =   1950
         Picture         =   "frmclients.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "SKIP"
         Top             =   3675
         Width           =   315
      End
      Begin VB.TextBox txttehseel 
         Height          =   330
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   22
         Top             =   3675
         Width           =   495
      End
      Begin VB.TextBox txtcityname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         Tag             =   "SKIP"
         Top             =   3300
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
         Left            =   1935
         Picture         =   "frmclients.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "SKIP"
         Top             =   3300
         Width           =   315
      End
      Begin VB.TextBox txtCity 
         Height          =   330
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   19
         Top             =   3300
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtpagrdate 
         Height          =   315
         Left            =   4965
         TabIndex        =   6
         Top             =   2145
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102694913
         CurrentDate     =   37848
      End
      Begin VB.TextBox txtmobile 
         Height          =   315
         Left            =   1425
         MaxLength       =   25
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1770
         Width           =   2400
      End
      Begin VB.TextBox txtphoneres 
         Height          =   315
         Left            =   4950
         MaxLength       =   25
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1770
         Width           =   2415
      End
      Begin VB.TextBox txtphoneoffice 
         Height          =   315
         Left            =   4935
         MaxLength       =   25
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2430
      End
      Begin VB.TextBox txtnic 
         Height          =   315
         Left            =   1425
         MaxLength       =   25
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1380
         Width           =   2400
      End
      Begin VB.TextBox txtaddress 
         Height          =   315
         Left            =   1425
         MaxLength       =   255
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   615
         Width           =   5940
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   0
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
         Left            =   2205
         Picture         =   "frmclients.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Disc Per % :"
         Height          =   210
         Left            =   2340
         TabIndex        =   52
         Top             =   5280
         Width           =   885
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bill Copy :"
         Height          =   210
         Left            =   585
         TabIndex        =   50
         Top             =   5280
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Address :"
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   48
         Top             =   1005
         Width           =   1350
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Website :"
         Height          =   210
         Left            =   4290
         TabIndex        =   46
         Top             =   4845
         Width           =   675
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email :"
         Height          =   210
         Left            =   945
         TabIndex        =   42
         Top             =   4830
         Width           =   450
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sector Code :"
         Height          =   210
         Left            =   390
         TabIndex        =   40
         Top             =   2580
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "St.Reg# :"
         Height          =   195
         Left            =   420
         TabIndex        =   35
         Top             =   4470
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Gl Account# :"
         Height          =   210
         Left            =   420
         TabIndex        =   32
         Top             =   4080
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Country :"
         Height          =   210
         Left            =   735
         TabIndex        =   29
         Top             =   2970
         Width           =   660
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tehseel :"
         Height          =   210
         Left            =   750
         TabIndex        =   26
         Top             =   3720
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "City :"
         Height          =   210
         Left            =   1005
         TabIndex        =   25
         Top             =   3345
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contact Person :"
         Height          =   210
         Index           =   8
         Left            =   195
         TabIndex        =   18
         Top             =   2175
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Opening Date :"
         Height          =   210
         Index           =   7
         Left            =   3870
         TabIndex        =   17
         Top             =   2175
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mobile :"
         Height          =   210
         Index           =   6
         Left            =   840
         TabIndex        =   16
         Top             =   1785
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fax # :"
         Height          =   210
         Index           =   5
         Left            =   4425
         TabIndex        =   15
         Top             =   1815
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Phone Office :"
         Height          =   210
         Index           =   4
         Left            =   3885
         TabIndex        =   14
         Top             =   1425
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NTN# :"
         Height          =   210
         Index           =   3
         Left            =   915
         TabIndex        =   13
         Top             =   1425
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         Height          =   210
         Index           =   2
         Left            =   660
         TabIndex        =   12
         Top             =   645
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   210
         Index           =   0
         Left            =   2580
         TabIndex        =   11
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   915
         TabIndex        =   9
         Top             =   270
         Width           =   465
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   1058
      ButtonWidth     =   1402
      ButtonHeight    =   1005
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
               Picture         =   "frmclients.frx":0BB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":100A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":145E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":18B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":1D06
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":215A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclients.frx":28AE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu File_menu 
      Caption         =   "File"
      Begin VB.Menu New_Record 
         Caption         =   "New Record"
         Shortcut        =   ^N
      End
      Begin VB.Menu Edit_Record 
         Caption         =   "Edit Record"
         Shortcut        =   ^E
      End
      Begin VB.Menu Delete_Record 
         Caption         =   "Delete Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu Save_Record 
         Caption         =   "Save Record"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkSupp As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Clients As New Recordset
Dim PR_Cities As New Recordset
Dim PR_Country As New Recordset
Dim PR_GlSub2 As New Recordset
Dim PR_Tehseels As New Recordset
Dim PR_Sector As New Recordset
Dim pr_dumy As New Recordset
Dim PR_DumyBSetup As New Recordset
Dim pr_dumy1 As New Recordset
Dim lb_found As Boolean
Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCity
    Set PO_DESC = txtcityname
    Gs_SQL = "Select CityCode, CityName from Cities "
    Gs_FindFld = "CityName"
    Gs_OrderBy = "Order by CityName"
    MyLookupOLDB.Caption = "City"
    MyLookupOLDB.Show 1
    If txtCity <> "" Then Call txtCity_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcountry
    Set PO_DESC = txtcountryName
    Gs_SQL = "Select CountryCode, CountryName from Country "
    Gs_FindFld = "CountryName"
    Gs_OrderBy = "Order by CountryName"
    MyLookupOLDB.Caption = "Country"
    MyLookupOLDB.Show 1
    If txtcountry <> "" Then Call txtcountry_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Delete_record_Click()
       Mode = DentMode(Mode, 3, PR_Clients, frmClients, txtClientCode, txtdesc, Para_Rs, "IC_PartyCnt", 6, "ClientCode", "Description", 1, False, Toolbar1)
       cmdLookup.Enabled = True
       txtClientCode.Enabled = True
       txtsub0 = Clientcode + Right(txtClientCode, 4)
       txtClientCode.SetFocus
End Sub

Private Sub Edit_record_Click()
       Mode = DentMode(Mode, 2, PR_Clients, frmClients, txtClientCode, txtdesc, Para_Rs, "IC_PartyCnt", 6, "ClientCode", "Description", 1, False, Toolbar1)
       cmdLookup.Enabled = True
       txtClientCode.Enabled = True
       txtsub0 = Clientcode + Right(txtClientCode, 4)
       txtClientCode.SetFocus
End Sub

Private Sub New_Record_Click()
       Mode = DentMode(Mode, 1, PR_Clients, frmClients, txtClientCode, txtdesc, Para_Rs, "IC_PartyCnt", 6, "ClientCode", "Description", 1, False, Toolbar1)
       cmdLookup.Enabled = False
       txtClientCode.Enabled = False
       txtClientCode = maxtranscode
       txtsub0 = Clientcode + Right(txtClientCode, 4)
       txtdesc.SetFocus
End Sub

Private Sub Save_Record_Click()
       Mode = DentMode(Mode, 4, PR_Clients, frmClients, txtClientCode, txtdesc, Para_Rs, "IC_PartyCnt", 6, "ClientCode", "Description", 1, False, Toolbar1)
End Sub

Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtCity <> "" Then
      txtCity.Text = DoPad(txtCity.Text, txtCity.MaxLength)
      PR_Cities.Open "Select * from cities where citycode = '" & txtCity & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1

      If Not PR_Cities.EOF Then
           txtcityname = Trim(PR_Cities("Cityname") & "")
           If Mode = "A" Then
'                txtsub0 = Trim(PR_Cities("GLClient") & "") + Right(Trim(txtClientCode), 4)
'                txtsub0.Enabled = False
           End If
           If txttehseel.Enabled Then txttehseel.SetFocus
        Else
           Call SetErr(Gs_RecNFMsg, vbCritical)
           txtCity.SetFocus
      End If
     PR_Cities.Close
End If
End Sub

Private Sub txtContactperson_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpagrdate.SetFocus
End Sub

Private Sub txtcountry_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtcountry <> "" Then
      txtcountry.Text = DoPad(txtcountry.Text, txtcountry.MaxLength)
      PR_Country.Open "Select * from Country where Countrycode = '" & txtcountry & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1

      If Not PR_Country.EOF Then
           txtcountryName = PR_Country("Countryname")
           If txtCity.Enabled Then txtCity.SetFocus
        Else
           Call SetErr(Gs_RecNFMsg, vbCritical)
           txtcountry.SetFocus
      End If
     PR_Country.Close
End If
End Sub


Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtClientCode
    Set PO_DESC = txtdesc
    Gs_SQL = "Select ClientCode, Description from IC_Clients "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Supplier"
    MyLookupOLDB.Show 1

    If Len(txtClientCode) > 0 Then txtClientCode_KeyDown vbKeyReturn, vbKeyShift

End Sub
Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttehseel
    Set PO_DESC = txttehseelname
    Gs_SQL = "Select tehseelCode, tehseelName from tehseels "
    Gs_FindFld = "tehseelName"
    Gs_OtherPara = "Where Citycode = '" & txtCity & "'"
    Gs_OrderBy = "Order by tehseelName"
    MyLookupOLDB.Caption = "Tehseel"
    MyLookupOLDB.Show 1

    If Len(txttehseel) > 0 Then txttehseel_KeyDown vbKeyReturn, vbKeyShift

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
    Gs_SQL = "Select SectorCode, SectorDesc from Ic_Sectors "
    Gs_FindFld = "SectorDesc"
    Gs_OrderBy = "Order by SectorDesc"
    MyLookupOLDB.Caption = "Sector"
    MyLookupOLDB.Show 1

    If Len(txtsector) > 0 Then txtsector_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub txtdeliveryAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtnic.SetFocus
End Sub

Private Sub txtDesc_Change()
If txtdesc <> "" Then
txtSubDesc = txtdesc
End If
End Sub

Private Sub txtDesc_LostFocus()
txtdesc = UCase(txtdesc)
txtsector = "001"
txtcountry = "001"
txtCity = "001"
txttehseel = "001"
txtwebsite = "NA"
txtemail = "NA"
txtstreg = "NA"
If txtsector <> "" Then Call txtsector_KeyDown(vbKeyReturn, vbKeyShift)
If txtcountry <> "" Then Call txtcountry_KeyDown(vbKeyReturn, vbKeyShift)
If txtCity <> "" Then Call txtCity_KeyDown(vbKeyReturn, vbKeyShift)
If txttehseel <> "" Then Call txttehseel_KeyDown(vbKeyReturn, vbKeyShift)
txtaddress.SetFocus
End Sub

Private Sub txtsector_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtsector <> "" Then
      txtsector.Text = DoPad(txtsector.Text, txtsector.MaxLength)
      PR_Sector.Open "Select * from Ic_Sectors where Sectorcode = '" & txtsector & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
      If Not PR_Sector.EOF Then
           txtsectordesc = PR_Sector("SectorDesc")
           If txtcountry.Enabled Then txtcountry.SetFocus
        Else
           Call SetErr(Gs_RecNFMsg, vbCritical)
           txtsector.SetFocus
      End If
     PR_Sector.Close
End If
End Sub
Private Sub txtstreg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtemail.SetFocus
End Sub

Private Sub txtsub0_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Val(txtsub0.Text) <> 0 Then
      PR_GlSub2.Open "Select * from Gl_Detail where Accountno = '" & txtsub0.Text & "' and CompCode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_GlSub2.EOF Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             If txtsub0.Enabled Then txtsub0.SetFocus
             txtSubDesc.Text = ""
         Else
             txtSubDesc.Text = PR_GlSub2("Acct_Desc")
             If txtstreg.Enabled Then txtstreg.SetFocus
         End If
 PR_GlSub2.Close
End If
End Sub

Private Sub txttehseel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txttehseel <> "" Then
      txttehseel.Text = DoPad(txttehseel.Text, txttehseel.MaxLength)
      PR_Tehseels.Open "Select * from tehseels where citycode = '" & txtCity & "' and TehseelCode = '" & txttehseel.Text & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1

      If Not PR_Tehseels.EOF Then
           txttehseelname = PR_Tehseels("TehseelName")
           If txtsub0.Enabled Then
            txtsub0.SetFocus
           Else
            If txtstreg.Enabled Then txtstreg.SetFocus
           End If
        Else
           Call SetErr(Gs_RecNFMsg, vbCritical)
           txttehseel.SetFocus
      End If
     PR_Tehseels.Close
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Clients, frmClients, txtClientCode, txtdesc, Para_Rs, "IC_PartyCnt", 6, "ClientCode", "Description", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SALCLIEN01")
  SetToolBar(2) = chkRights("SALCLIEN02")
  SetToolBar(3) = chkRights("SALCLIEN03")
  SetToolBar(4) = chkRights("SALCLIEN04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  txtType = "Customer"
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtaddress.SetFocus
    If KeyCode = vbKeyPageUp And txtClientCode.Enabled Then txtClientCode.SetFocus
End Sub
Private Sub txtaddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtdeliveryAddress.SetFocus
    If KeyCode = vbKeyPageUp Then txtdesc.SetFocus
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
    If KeyCode = vbKeyReturn Then txtcontactperson.SetFocus
    If KeyCode = vbKeyPageUp Then txtmobile.SetFocus
End Sub
Private Sub txtmobile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtphoneres.SetFocus
    If KeyCode = vbKeyPageUp Then txtphoneoffice.SetFocus
End Sub
Private Sub dtpagrdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtsector.SetFocus
    If KeyCode = vbKeyPageUp Then txtType.SetFocus
End Sub
Private Sub txtClientCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Trim(txtClientCode.Text) <> "" Then
       txtClientCode.Text = DoPad(UCase(txtClientCode.Text), txtClientCode.MaxLength)
       If PR_Clients.State = 1 Then PR_Clients.Close
       PR_Clients.Open "Select * from IC_Clients where Compcode  = '" & Gs_compcode & "' and ClientCode = '" & txtClientCode.Text & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       Select Case Mode
            Case "A"
                If Not PR_Clients.EOF Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   txtClientCode.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If PR_Clients.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Call ClearVal
                   txtClientCode.SetFocus
                Else
                   Call SetVal
                End If
     PR_Clients.Close
     End Select
ElseIf KeyCode = vbKeyReturn And Trim(txtClientCode.Text) = "" Then
    txtClientCode = ""
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
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Clients, frmClients, txtClientCode, txtdesc, Para_Rs, "IC_PartyCnt", 6, "ClientCode", "Description", 1, False, Toolbar1)
    End If
    
    If Mode = "A" Then
       txtClientCode = maxtranscode
       txtsub0 = Clientcode + Right(txtClientCode, 4)
       txtdesc.SetFocus
    End If

End Sub

Public Sub SaveValues()
Dim ln_cnt As Integer
Dim ls_CodeID As String



gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT into IC_Clients(Compcode,ClientCode,Description,Address,NTNNo,Phoneoffice,Phoneres,Mobile,AgrDate,Contactperson,CountryCode,CityCode,TehseelCode,GlAccountNo,stregno,sectorcode,email,Website,DeliveryAddress,BillCopy,DiscAllowYn,Discper) VALUES ('" & Gs_compcode & "','" & txtClientCode.Text & "','" & txtdesc.Text & "','" & txtaddress.Text & "','" & txtnic.Text & "','" & txtphoneoffice.Text & "','" & txtphoneres.Text & "','" & txtmobile.Text & "','" & Format(dtpagrdate, "YYYY/MM/DD") & "','" & txtcontactperson & "','" & Trim(txtcountry) & "','" & Trim(txtCity) & "','" & Trim(txttehseel) & "','" & Trim(txtsub0) & "','" & txtstreg & "','" & txtsector & "','" & txtemail & "','" & txtwebsite & "','" & txtdeliveryAddress & "'," & Val(txtBillCopy) & "," & Val(txtdiscper) & "," & Val(ChkDiscAllow.Value) & ")"
              cmdLookup.Enabled = False
              
           Case "E"
              gc_dbcon.Execute "UPDATE IC_Clients SET Description= '" & txtdesc.Text & "',Address= '" & txtaddress.Text & "',NtnNo= '" & txtnic.Text & "',Phoneoffice= '" & txtphoneoffice.Text & "',PhoneRes= '" & txtphoneres.Text & "',Mobile= '" & txtmobile.Text & "',AgrDate= '" & Format(dtpagrdate, "YYYY/MM/DD") & "', ContactPerson ='" & txtcontactperson & "', countryCode  = '" & txtcountry & "', CityCode  = '" & txtCity & "', TehseelCode  = '" & txttehseel & "', GlAccountNo  = '" & Trim(txtsub0) & "',stregno = '" & txtstreg & "',sectorcode = '" & txtsector & "',email = '" & Trim(txtemail) & "',website = '" & Trim(txtwebsite) & "',DeliveryAddress = '" & Trim(txtdeliveryAddress) & "',BillCopy = " & Val(txtBillCopy) & ",DiscPer =  " & Val(txtdiscper) & ", DiscAllowYN =  " & Val(ChkDiscAllow.Value) & " WHERE  ClientCode= '" & txtClientCode.Text & "' and Compcode = '" & Gs_compcode & "'"
              
           Case "D"
              gc_dbcon.Execute "DELETE FROM IC_Clients WHERE ClientCode = '" & txtClientCode.Text & "' and Compcode = '" & Gs_compcode & "'"
           
     End Select
gc_dbcon.CommitTrans

'glCode

pr_dumy.Open "Select * from Gl_Detail where compcode = '" & Gs_compcode & "' and accountno = '" & txtsub0 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If pr_dumy.EOF Then

ls_sql = "INSERT into Gl_detail(compcode,Acct_sub,Acct_Detail,AccountNo,Acct_desc,crncy_code,Acct_Base,Acct_Type,Acct_Status,userid,adddate,addtime) VALUES ('" & Gs_compcode & "','" & Left(txtsub0, 9) & "', '" & Right(txtClientCode, 4) & " ','" & txtsub0 & "','" & txtdesc & "','PKR','B','G','D','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
gc_dbcon.Execute ls_sql

Else
    If Mode = "A" Or Mode = "E" Then
        Call MsgBox("GL Code for this Client already exist - [" & Trim(pr_dumy("acct_desc") & "") & "]" & Chr(13) & "System has saved the same code for Sale transaction " & Chr(13) & "User can change this code by edit and select the same code as in GL ", vbInformation)
    End If
End If
pr_dumy.Close

'end gl Code

'Balance Sheet





'PR_Dumy.Open "Select * from GL_Bsheet3Detail where compcode = '" & Gs_compcode & "' and accountno = '" & txtsub0 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'If PR_Dumy.EOF Then
'
'
'
'
'    PR_DumyBSetup.Open "Select * from VendorBSSetup where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'    If Not PR_DumyBSetup.EOF Then
'    ls_bcode = Trim(PR_DumyBSetup("BSMainHead") & "")
'    ls_bncode = Trim(PR_DumyBSetup("BSNotes") & "")
'    End If
'    PR_DumyBSetup.Close
'
'    ls_bnsicode = maxBSCode(Gs_compcode)
'
'    If ls_bcode <> "" Then
'        ls_sql = "insert into  gl_bsheet3 (compcode,bcode,bncode,bnicode,bnidesc) values ('" & Gs_compcode & "','" & ls_bcode & "','" & ls_bncode & "','" & ls_bnsicode & "','" & txtDesc & "')"
'        gc_dbcon.Execute ls_sql
'
'        ls_sql = "insert into  gl_bsheet3detail (compcode,bcode,bncode,bnicode,accountno) values ('" & Gs_compcode & "','" & ls_bcode & "','" & ls_bncode & "','" & ls_bnsicode & "','" & txtsub0 & "')"
'        gc_dbcon.Execute ls_sql
'    Else
'       Call MsgBox("Vendor Balance Sheet not setup", vbCritical)
'    End If
'
'End If
'PR_Dumy.Close
'
''end balance sheet


'opening balance adjustment


'If Mode = "A" Or Mode = "E" Then
'Dim ls_transcode As String
'Dim ls_Referencecode As String
'ls_Referencecode = "S" + txtClientCode
'ls_transcode = maxtranscodepayments
'gc_dbcon.Execute "Delete from Ic_payments where compcode = '" & Gs_compcode & "' and referencecode = '" & ls_Referencecode & "' and codeid = 'S' "
'    If Val(txtopeningbal) > 0 Then
'        ls_sql = "Insert into IC_Payments(Compcode,PartyCode,codeid,TransCode,TransDate,Amount,taxamount,Remarks,ReferenceCode) Values ('" & Gs_compcode & "','" & txtClientCode & "','" & Left(txtType, 1) & "','" & ls_transcode & "','" & Format(dtpagrdate, "YYYY/MM/DD") & "' ," & Val(txtopeningbal) & ",0,'Opening Balance','" & ls_Referencecode & "' )"
'        gc_dbcon.Execute ls_sql
'    End If
'End If
'
If Mode = "A" Then
       txtClientCode = maxtranscode
       txtdesc.SetFocus
End If


Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Function maxBSCode(cmpCode As String) As String


PR_DumyBSetup.Open "Select * from VendorBSSetup where compcode = '" & cmpCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
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

Private Function maxtranscodepayments() As String
Dim pr_dumypmt As New Recordset
pr_dumypmt.Open "select max(transcode) as transcode from ic_payments where compcode = '" & Gs_compcode & "' and codeid = 'S'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumypmt.EOF Then
    maxtranscodepayments = DoPad(Trim(str(Val(0 & pr_dumypmt("transcode")) + 1)), 10)
Else
    maxtranscodepayments = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumypmt.Close
End Function

Private Function maxtranscode() As String
pr_dumy.Open "select max(ClientCode) as transcode from IC_Clients where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 6)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 6)
End If
pr_dumy.Close
End Function



Public Sub ClearVal()
     txtClientCode = ""
     txtdesc = ""
     txtCrValue = ""
     txtCrQty = ""
     'Option1(0).Value = True
End Sub

Private Sub SetVal()
     Dim ln_cnt As Integer
     txtdesc = PR_Clients("Description")
     txtaddress = Trim(PR_Clients("Address") & "")
     txtdeliveryAddress = Trim(PR_Clients("DeliveryAddress") & "")
     txtnic = Trim(PR_Clients("NTNNo") & "")
     txtphoneoffice = Trim(PR_Clients("Phoneoffice") & "")
     txtphoneres = Trim(PR_Clients("PhoneRes") & "")
     txtmobile = Trim(PR_Clients("Mobile") & "")
     dtpagrdate = PR_Clients("agrdate")
     txtcountry = PR_Clients("CountryCode") & ""
     txtcontactperson = PR_Clients("ContactPerson") & ""
     txtstreg = PR_Clients("stregno") & ""
     txtemail = Trim(PR_Clients("email") & "")
     
    ' txtBillCopy = Val(PR_Clients("BillCopy"))
    ' txtdiscper = Val(PR_Clients("Discper"))
    ' ChkDiscAllow.Value = Val(PR_Clients("DiscAllowYN"))
     
     
     
     'txtopeningbal = Val(PR_Clients("openingbalance"))
     If txtcountry <> "" Then Call txtcountry_KeyDown(vbKeyReturn, vbKeyShift)
     txtCity = PR_Clients("CityCode") & ""
     If txtCity <> "" Then Call txtCity_KeyDown(vbKeyReturn, vbKeyShift)
     txtsub0 = Trim(PR_Clients("GlAccountNo") & "")
     If txtsub0 <> "" Then Call txtsub0_KeyDown(vbKeyReturn, vbKeyShift)
     txtsector = PR_Clients("SectorCode") & ""
     If txtsector <> "" Then Call txtsector_KeyDown(vbKeyReturn, vbKeyShift)
     txttehseel = PR_Clients("TehseelCode") & ""
     If txttehseel <> "" Then Call txttehseel_KeyDown(vbKeyReturn, vbKeyShift)

End Sub
Public Function ChkInputs() As Boolean
    If Len(txtClientCode.Text) = txtClientCode.MaxLength And txtdesc.Text <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtdesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub
