VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPRAuth_GRN 
   Caption         =   "Authentication & GRN"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPRAuth_GRN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7110
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
      Height          =   1980
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   7095
      Begin VB.TextBox txtInPass 
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
         MaxLength       =   5
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   525
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   2520
         Picture         =   "FrmPRAuth_GRN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   525
         Width           =   315
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   4920
         TabIndex        =   32
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   37410
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3060
         MaxLength       =   50
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   315
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1590
         Width           =   4695
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   2520
         Picture         =   "FrmPRAuth_GRN.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1230
         Width           =   315
      End
      Begin VB.TextBox TxtPartyCode 
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1095
      End
      Begin VB.TextBox TxtPartyDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2940
         MaxLength       =   64
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1230
         Width           =   3195
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2520
         Picture         =   "FrmPRAuth_GRN.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   870
         Width           =   315
      End
      Begin VB.TextBox txtLocCode 
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   870
         Width           =   1095
      End
      Begin VB.TextBox txtLocDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2940
         MaxLength       =   64
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   870
         Width           =   3195
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2535
         Picture         =   "FrmPRAuth_GRN.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
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
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Inword Pass # :"
         Height          =   255
         Left            =   60
         TabIndex        =   35
         Top             =   525
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1590
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Supplier Id :"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1230
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Location In :"
         Height          =   255
         Left            =   60
         TabIndex        =   24
         Top             =   870
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Reference #  :"
         Height          =   255
         Left            =   90
         TabIndex        =   23
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label label2 
         Caption         =   "Value Date :"
         Height          =   255
         Left            =   4020
         TabIndex        =   22
         ToolTipText     =   "Enter Value Date"
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2475
      Left            =   0
      TabIndex        =   1
      Top             =   2460
      Width           =   7110
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   5535
         MaxLength       =   11
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Quantity"
         Top             =   420
         Width           =   780
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   3630
         TabIndex        =   37
         Top             =   420
         Width           =   945
      End
      Begin VB.TextBox txtItemDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   60
         MaxLength       =   64
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2115
         Width           =   2835
      End
      Begin VB.TextBox TxtGrnTotal 
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
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4005
         MaxLength       =   11
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Total GRN Value"
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox TxtUM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0;(""$""#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2895
         MaxLength       =   11
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   420
         Width           =   720
      End
      Begin VB.TextBox TxtFactor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0;(""$""#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2475
         MaxLength       =   11
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   420
         Width           =   435
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2190
         Picture         =   "FrmPRAuth_GRN.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   435
         Width           =   315
      End
      Begin VB.TextBox TxtItemCode 
         BackColor       =   &H00FFFFFF&
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
         Left            =   45
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Item Code"
         Top             =   435
         Width           =   2115
      End
      Begin VB.TextBox txtqty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4590
         MaxLength       =   11
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Quantity"
         Top             =   420
         Width           =   930
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.000;(""$""#,##0.000)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   315
         Left            =   6330
         MaxLength       =   11
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Unit Price"
         Top             =   420
         Width           =   720
      End
      Begin MSFlexGridLib.MSFlexGrid GrdGRN 
         Height          =   1335
         Left            =   30
         TabIndex        =   18
         Top             =   720
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   1
         Cols            =   8
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Rej. Qty :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5610
         TabIndex        =   40
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Description :"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3690
         TabIndex        =   38
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label11 
         Caption         =   "GRN. Total :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3045
         TabIndex        =   29
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label10 
         Caption         =   "U / M :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3105
         TabIndex        =   28
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Factor :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2475
         TabIndex        =   26
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Item Code :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   105
         TabIndex        =   21
         Top             =   180
         Width           =   2205
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Quantity :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4695
         TabIndex        =   20
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Unit Price:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6330
         TabIndex        =   19
         Top             =   180
         Width           =   720
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   7110
      _ExtentX        =   12541
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
            Caption         =   "&Print"
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
         Left            =   4995
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
               Picture         =   "FrmPRAuth_GRN.frx":0A44
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPRAuth_GRN.frx":0E98
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPRAuth_GRN.frx":12EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPRAuth_GRN.frx":1740
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPRAuth_GRN.frx":1B94
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPRAuth_GRN.frx":1FE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPRAuth_GRN.frx":273C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmPRAuth_GRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PB_BlnkGRN As Boolean
Dim Mode As String

Public PO_CODE As Object
Public PO_DESC As Object
'Public PL_Status As Boolean

Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String

Dim cntsql As New ADODB.Command
Dim ls_TransType As String

Dim PR_PRGRNSUM As New Recordset
Dim PR_PRItmLoc As New Recordset
Dim PR_InTrans As New Recordset
Dim PR_PRParty As New Recordset
Dim PR_PRGRN As New Recordset
Dim PR_PRItem As New Recordset
Dim PR_InPass As New Recordset
Dim PR_GRN_No As New Recordset

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtTransNo
    Set PO_DESC = Text1
    GoTop PR_PRGRN
    MyLookup.Caption = "GRNs "
    MyLookup.FillGrid PR_PRGRN, "GrnNo", "Value_Date", 10
    MyLookup.Show 1
    
    If Len(txtTransNo) > 0 Then txtTransNo_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = txtLocDesc
    GoTop PR_PRItmLoc
    MyLookup.Caption = "Items. "
    MyLookup.FillGrid PR_PRItmLoc, "LocationCode", "Description", 6
    MyLookup.Show 1
    
    If Len(txtLocCode) > 0 Then txtLocCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command3_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtPartyCode
    Set PO_DESC = TxtPartyDesc
    GoTop PR_PRParty
    MyLookup.Caption = "Supplier Parties. "
    MyLookup.FillGrid PR_PRParty, "supplierCode", "Description", 6
    MyLookup.Show 1
    
    If Len(TxtPartyCode) > 0 Then txtPartyCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command1_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtItemCode
    Set PO_DESC = txtItemDesc
    
    PR_InPass.Filter = "InPassNo = '" & txtInPass.Text & "'"
    If PR_InPass.EOF Then
       Call SetErr("No Item has been found.", vbCritical)
       Exit Sub
    End If
    GoTop PR_InPass
    MyLookup.Caption = "Items. "
    MyLookup.FillGrid PR_InPass, "ItemCode", "Description", 25
    MyLookup.Show 1
    PR_InPass.Filter = adFilterNone
    
    If Len(TxtItemCode) > 0 Then txtItemCode_KeyDown vbKeyReturn, vbKeyShift
  
End Sub

Private Sub Command4_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtInPass
    Set PO_DESC = Text1
    GoTop PR_InPass
    lb_found = MySeek(txtInPass, "InPassNo", PR_InPass)

    MyLookup.Caption = "Purchase Order Nos "
    MyLookup.FillGrid PR_InPass, "InPassNo", "Value_Date", 5
    MyLookup.Show 1
    
    If Len(txtInPass) > 0 Then txtInPass_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_Load()
 cntsql.ActiveConnection = gc_dbcon
 cntsql.CommandType = adCmdText
 ls_TransType = "G"
  SetToolBar(1) = chkRights("ICGRNSTP01")
  SetToolBar(2) = False
  SetToolBar(3) = False
  SetToolBar(4) = chkRights("ICGRNSTP04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  gc_dbcon.Execute ("Select SUM(Round(GrnQty*UnitPrice,0)) As TGRNValue, SUM(GrnQTY) AS TGRNQty,ItemCode as ItemGRN into Tmp_GRNSUM from PR_GRN  where compcode ='" & Gs_compcode & "' AND Value_Date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' AND Value_Date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "' group by itemcode")
  
  PR_GRN_No.Open "Select Distinct(PR_GRN.GrnNo) as GrnNo,Value_Date as Value_Date From PR_GRN ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_PRItmLoc.Open "Select * from Ic_Locations where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_PRParty.Open "Select * from Ic_Supplier ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_PRItem.Open "Select *,(ltrim(rtrim(itemclass))+ltrim(rtrim(itemcode))) as ItemID,(ltrim(rtrim(locationcode))+ltrim(rtrim(itemclass))+ltrim(rtrim(itemcode))) as ItemFind from Ic_Item where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_PRGRNSUM.Open "Select Tmp_GRNSUM.* From Tmp_GRNSUM Order by ItemGRN", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_PRGRN.Open "Select *,(Compcode + GrnNo) As GRNFind ,(Compcode + GrnNo + ItemCode) As GRNItemFind from PR_Grn where compcode ='" & Gs_compcode & "' AND (Value_Date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' AND Value_Date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "') order by grnfind", gc_dbcon, adOpenDynamic, adLockOptimistic
  PR_InPass.Open "Select PR_InWordPass.*,IC_Item.Description as Description,IC_Item.ItemUM as ItemUM,IC_Item.ItemFactor as ItemFactor,IC_Item.ItemUnitPrice as PurchaseRate,(LTrim(RTrim(PR_InWordPass.InPassNo)) + LTrim(RTrim(PR_InWordPass.ItemCode))) as ItemFind From PR_InWordPass,IC_Item Where PR_InWordPass.CompCode = '" & Gs_compcode & "' and RTrim(LTrim(PR_InWordPass.ItemCode)) = RTrim(LTrim(IC_Item.ItemID)) order by PR_InWordPass.InPassNo ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  
  PB_BlnkGRN = IIf(PR_PRGRN.EOF, True, False)
  txtvaluedate.Value = Date
  InitializeGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_PRItmLoc.Close
    PR_PRParty.Close
    PR_PRGRN.Close
    PR_PRItem.Close
    Set PR_PRGRNSUM = Nothing
    gc_dbcon.Execute ("Drop Table Tmp_GRNSUM;")
End Sub

Private Sub Text1_GotFocus()
    txtvaluedate.SetFocus
End Sub

Private Sub txtInPass_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

If LastKey(KeyCode) And Len(txtInPass.Text) > 0 Then
        
    txtInPass.Text = IIf(IsNumeric(LTrim(Str(txtInPass.Text))), DoPad(UCase(txtInPass.Text), txtInPass.MaxLength), UCase(txtInPass.Text))
    lb_found = MySeek(txtInPass, "InPassNo", PR_InPass)
    If Not lb_found Then
        Call SetErr(Gs_RecNFMsg, vbCritical)
        txtInPass.SetFocus
    Else
        txtLocCode.SetFocus
    End If
End If
End Sub

Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If Val(txtqty.Text) <= 0 Then
          Call SetErr("Invalid Quantity Value.", vbCritical)
          txtqty.SetFocus
       Else
          txtqty.Text = Round(Val(txtqty.Text) * Val(TxtFactor), 0)
          txtUnitPrice.SetFocus
       End If
    End If
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
   If LastKey(KeyCode) Then
      TxtItemCode.SetFocus
   End If
End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If LastKey(KeyCode) And Len(txtTransNo.Text) > 0 Then
         
         txtTransNo.Text = IIf(IsNumeric(LTrim(Str(txtTransNo.Text))), DoPad(UCase(txtTransNo.Text), 10), UCase(txtTransNo.Text))
         lb_found = MySeek(Gs_compcode & txtTransNo.Text, "GRNFind", PR_PRGRN)
         
         If Gl_Demo Then
            If PR_PRGRN.RecordCount > gn_MaxVchrs Then
                Call SetErr("This is Demo software you cannot add more than " & LTrim(Str(gn_MaxVchrs)) & " Transactions.", vbCritical)
                Unload Me
                Exit Sub
            End If
         End If

         InitializeGrid
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   Cancel = True
                   Call ClearVal
                   txtTransNo.SetFocus
                Else
                   txtLocCode.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                   Call ClearVal
                   txtLocCode.SetFocus
                Else
                   Call SetVal
'                   LoadGRNTrans
                   If Mode <> "D" Then
                      txtTransNo.SetFocus
                   End If
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

    If PB_BlnkGRN And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
'      Mode = DentMode(Mode, Button.Index, PR_PRGRN, Me, txtTransNo, txtvaluedate, "X", "IC_GRNCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)
      Mode = DentMode(Mode, Button.Index, PR_PRGRN, Me, txtTransNo, txtvaluedate, Para_Rs, "IC_GRNCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)

    End If
    
End Sub


Public Sub SaveValues()
Dim cntsql As New ADODB.Command
Dim Ln_Cnt As Integer
Dim ln_AvgPrice As Integer
Dim ln_OpnQty As Double
Dim ln_OpnValu As Double

On Error GoTo RollBack
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

If Mode = "A" Then FrmRelGRN.Show 1

gc_dbcon.BeginTrans
     Select Case Mode
           Case "D"
'              cntsql.CommandText = "DELETE FROM IC_Trans WHERE CompCode = '" & Gs_compcode & "' AND Transc_No = '" & Trim(txtTransNo) & "' AND TransType = '" & ls_TransType & "' AND Value_Date = '" & Format(txtvaluedate, "YYYY/MM/DD") & "' AND LocationCode1 = '" & Trim(txtLocCode) & "'"
'              cntsql.Execute
           Case Else
'                If Mode = "E" Then
'                    cntsql.CommandText = "DELETE FROM IC_Trans WHERE CompCode = '" & Gs_compcode & "' AND Transc_No = '" & Trim(txtTransNo) & "' AND TransType = '" & ls_TransType & "' AND Value_Date = '" & Format(txtvaluedate, "YYYY/MM/DD") & "' AND RTrim(LTrim(LocationCode1)) = '" & Trim(txtLocCode) & "'"
'                    cntsql.Execute
'                End If
              
                With GrdGRN
                    For Ln_Cnt = 1 To .Rows - 1
                       ln_OpnQty = 0
                       ln_OpnValue = 0
                       If MySeek(txtLocCode.Text & .TextMatrix(Ln_Cnt, 1), "ItemGRN", PR_PRGRNSUM) Then
                          PR_PRGRNSUM.Fields("TGRNValue") = PR_PRGRNSUM.Fields("TGRNValue") + Val(.TextMatrix(Ln_Cnt, 4))
                          PR_PRGRNSUM.Fields("TGRNQty") = PR_PRGRNSUM.Fields("TGRNQty") + Val(.TextMatrix(Ln_Cnt, 2))
                          'PR_PRGRNSUM.Update
                       End If
                       
                       If MySeek(txtLocCode.Text & .TextMatrix(Ln_Cnt, 1), "ItemFind", PR_PRItem) Then
                          ln_OpnQty = Val(PR_PRItem.Fields("ItemOpenQty") & "") + 0
                          ln_OpnValue = Val(PR_PRItem.Fields("ItemOpenValue") & "") + 0
                          
                          If PR_PRGRNSUM.EOF Then
                             ln_AvgPrice = Round((ln_OpnValue + Val(.TextMatrix(Ln_Cnt, 4))) / (Val(.TextMatrix(Ln_Cnt, 2)) + ln_OpnQty), 2)
                          Else
                             ln_AvgPrice = Round((ln_OpnValue + PR_PRGRNSUM.Fields("TGRNValue")) / (ln_OpnQty + PR_PRGRNSUM.Fields("TGRNQty")), 2)
                          End If
                          PR_PRItem.Fields("ItemAvgPrice") = ln_AvgPrice
                          PR_PRItem.Fields("ItemUnitPrice") = Val(.TextMatrix(Ln_Cnt, 3))
                          PR_PRItem.Fields("ItemBalQty") = PR_PRItem.Fields("ItemBalQty") + Val(.TextMatrix(Ln_Cnt, 2))
                          PR_PRItem.Update
                       End If
                       
                       cntsql.CommandText = "INSERT into IC_Trans(compcode,Transc_No, TransType, Value_Date, LocationCode1, SerialNo, SupplierCode, ItemClass, ItemCode,Quantity,UnitCost,Remarks) VALUES ('" & Gs_compcode & "','" & Trim(txtTransNo) & "','" & ls_TransType & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & Trim(txtLocCode) & "'," & .TextMatrix(Ln_Cnt, 0) & ",'" & Trim(TxtPartyCode) & "','" & .TextMatrix(Ln_Cnt, 6) & "','" & .TextMatrix(Ln_Cnt, 7) & "'," & Val(.TextMatrix(Ln_Cnt, 2)) & "," & Val(.TextMatrix(Ln_Cnt, 3)) & ",'" & Trim(TxtRemarks) & "')"
                       cntsql.Execute
                       
                   Next
                 txtTransNo.Text = DoPad(LTrim(Str(Para_Rs.Fields("IC_GRNCnt") + 1)), 10)
                 End With
                 
     End Select
gc_dbcon.CommitTrans
PR_PRGRN.Requery
On Error GoTo 0
Exit Sub

RollBack:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
gc_dbcon.RollbackTrans

If Mode = "A" Then
    Para_Rs.Fields("Ic_GRNcnt") = Para_Rs.Fields("Ic_GRNcnt") - 1
    Para_Rs.Update
End If

On Error GoTo 0
End Sub
Public Sub ClearVal()
     txtLocDesc = ""
     TxtPartyDesc = ""
     txtLocCode = ""
     TxtPartyCode = ""
     TxtRemarks = ""
     txtvaluedate.Value = Date
     TxtGrnTotal = ""
     TxtItemCode = ""
     InitializeGrid
     TxtFactor = ""
     TxtUM = ""
     PI_SrNo = 0
End Sub

Private Sub SetVal()
     txtvaluedate = PR_PRGRN("Value_Date")
     txtLocCode = PR_PRGRN("LocationCode")
     txtLocDesc = IIf(MySeek(PR_PRGRN("LocationCode"), "LocationCode", PR_PRItmLoc), PR_PRItmLoc.Fields("Description"), "Not Found.")
     TxtPartyCode = PR_PRGRN("SupplierCode")
     TxtPartyDesc = IIf(MySeek(PR_PRGRN("SupplierCode"), "SupplierId", PR_PRParty), PR_PRParty.Fields("Description"), "Not Found.")
     TxtRemarks = PR_PRGRN("Remarks")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtTransNo.Text) = txtTransNo.MaxLength And Len(txtLocCode) = txtLocCode.MaxLength And Len(TxtPartyCode) = TxtPartyCode.MaxLength And PI_SrNo > 0 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If LastKey(KeyCode) And Len(txtLocCode.Text) > 0 Then
         txtLocCode.Text = IIf(IsNumeric(LTrim(RTrim(txtLocCode.Text))), DoPad(txtLocCode.Text, 6), UCase(txtLocCode.Text))
         lb_found = MySeek(txtLocCode.Text, "LocationCode", PR_PRItmLoc)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtLocCode.SetFocus
             txtLocDesc.Text = ""
         Else
             txtLocDesc.Text = PR_PRItmLoc("Description")
             TxtPartyCode.SetFocus
         End If
 End If
End Sub

Private Sub txtPartyCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If LastKey(KeyCode) And Len(TxtPartyCode.Text) > 0 Then
         TxtPartyCode.Text = IIf(IsNumeric(LTrim(RTrim(TxtPartyCode.Text))), DoPad(TxtPartyCode.Text, 6), UCase(TxtPartyCode.Text))
         lb_found = MySeek(TxtPartyCode.Text, "SupplierCode", PR_PRParty)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             TxtPartyCode.SetFocus
             TxtPartyDesc.Text = ""
         Else
             TxtPartyDesc.Text = PR_PRParty("Description")
             TxtRemarks.SetFocus
         End If
 End If
End Sub
Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
Dim Ln_Cnt As Integer
 
 If LastKey(KeyCode) And Len(TxtItemCode.Text) > 0 Then
         With GrdGRN
               For Ln_Cnt = 1 To .Rows - 1
                   If .TextMatrix(Ln_Cnt, 1) = TxtItemCode Then
                      Call SetErr("Item ID already exists.", vbCritical)
                      TxtItemCode.Text = ""
                      TxtItemCode.SetFocus
                      Exit Sub
                   End If
                Next
         End With
         TxtItemCode.Text = UCase(TxtItemCode.Text)
         lb_found = MySeek(txtLocCode.Text & TxtItemCode.Text, "ItemFind", PR_PRItem)
         
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             TxtItemCode.SetFocus
         Else
             txtItemDesc = PR_PRItem.Fields("Description")
             TxtFactor = PR_PRItem.Fields("ItemFactor")
             TxtUM = PR_PRItem.Fields("ItemUM")
             txtqty.SetFocus
         End If
 End If
End Sub

Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Item Code|Description|<Quantity|Rejected Qty|<Unit Price|<Amount  |<GLAccount"
        .ColWidth(1) = 2000
        .ColWidth(2) = 900
        .ColWidth(3) = 900
        .ColAlignment(3) = 7
        .ColWidth(4) = 1100
        .ColAlignment(4) = 7
        .ColWidth(5) = 750
        .ColWidth(6) = 900
        .ColWidth(7) = 0
        .Redraw = True
    End With
End Sub

Private Sub grdGrn_KeyDown(KeyCode As Integer, Shift As Integer)
    With GrdGRN
        If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            TotalGRN
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
            End If
        End If
    End With
End Sub

Private Sub txtUnitPrice_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Ln_Cnt As Integer

    If LastKey(KeyCode) Then
      If (Val(txtqty) > 0 Or Val(txtUnitPrice) > 0) Then
        
        If PS_RowClicked = "" Then
            If PI_SrNo = 0 Then
                PI_SrNo = 1
            Else
                PI_SrNo = PI_SrNo + 1
            End If
        End If
        
        If Len(TxtItemCode.Text) > 0 Then
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
              If MySeek(txtLocCode.Text & TxtItemCode.Text, "ItemFind", PR_PRItem) Then
                .TextMatrix(.Row, 1) = Trim(TxtItemCode)
                .TextMatrix(.Row, 2) = Val(txtqty)
                .TextMatrix(.Row, 3) = Val(txtUnitPrice)
                .TextMatrix(.Row, 4) = Round(Val(txtqty) * Val(txtUnitPrice), 0)
                
                .TextMatrix(.Row, 5) = PR_PRItem("LocationCode")
                .TextMatrix(.Row, 6) = PR_PRItem("ItemClass")
                .TextMatrix(.Row, 7) = PR_PRItem("ItemCode")
                .TextMatrix(.Row, 8) = PR_PRItem("Gl_AccountNo") & ""
                
                TxtItemCode.Text = ""
                TxtFactor.Text = ""
                TxtUM.Text = ""
                txtqty = ""
                txtUnitPrice = ""
                PS_RowClicked = ""
                TxtItemCode.SetFocus
             Else
                Call SetErr("Item Code Not Found.", vbCritical)
                TxtItemCode.SetFocus
             End If
            End With
        End If
            TotalGRN
    Else
               TxtItemCode.SetFocus
    End If
    End If
End Sub

Private Sub TotalGRN()
    Dim Ln_Cnt As Integer
    Dim Ln_Cnt2 As Integer
    Dim lb_AppArray As Integer
    Dim ln_ArRow As Integer
    Dim ls_xx As String
    TxtGrnTotal = ""
    ln_ArRow = 1
    
    With GrdGRN
        ReDim gv_glArray(1 To (.Rows - 1), 2)
        For Ln_Cnt = 1 To .Rows - 1
            .TextMatrix(Ln_Cnt, 0) = Ln_Cnt
            TxtGrnTotal = Val(TxtGrnTotal) + Val(.TextMatrix(Ln_Cnt, 4))
            PI_SrNo = Ln_Cnt
            
            'Search AccontNos in Array
            lb_AppArray = True
            For Ln_Cnt2 = 1 To .Rows - 1
                If .TextMatrix(Ln_Cnt, 8) = gv_glArray(Ln_Cnt2, 1) And .TextMatrix(Ln_Cnt2, 8) <> "" Then
                    lb_AppArray = False
                    Exit For
                End If
            Next
            ' Save into gv_GlArray
            If .TextMatrix(Ln_Cnt, 8) <> "" Then
            If lb_AppArray Then
               gv_glArray(ln_ArRow, 1) = .TextMatrix(Ln_Cnt, 8)
               gv_glArray(ln_ArRow, 2) = .TextMatrix(Ln_Cnt, 4)
               ln_ArRow = ln_ArRow + 1
            Else
               gv_glArray(Ln_Cnt2, 2) = gv_glArray(Ln_Cnt2, 2) + Val(.TextMatrix(Ln_Cnt, 4))
            End If
            End If
        Next
    End With
End Sub

Private Sub LoadGRNTrans()
Dim lb_found As Boolean
Dim Ln_Cnt   As Integer
Dim temp As String

InitializeGrid
    
    lb_found = MySeek(Gs_compcode & ls_TransType & txtTransNo, "Grnfind", PR_PRGRN)
   
    If lb_found Then
        With GrdGRN
            Do While Gs_compcode & ls_TransType & txtTransNo = PR_PRGRN("GRNFind")
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = PR_PRGRN("SerialNo")
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_PRGRN("LocationCode1")) & Trim(PR_PRGRN("ItemClass")) & Trim(PR_PRGRN("ItemCode"))
                .TextMatrix(.Row, 2) = PR_PRGRN("Quantity") & ""
                .TextMatrix(.Row, 3) = PR_PRGRN("UnitCost")
                .TextMatrix(.Row, 4) = Round(PR_PRGRN("Quantity") * PR_PRGRN("UnitCost"), 0)
                .TextMatrix(.Row, 5) = Trim(PR_PRGRN("LocationCode1"))
                .TextMatrix(.Row, 6) = PR_PRGRN("ItemClass")
                .TextMatrix(.Row, 7) = PR_PRGRN("ItemCode")
'                .TextMatrix(.Row, 8) = PR_PRGRN("Gl_AccountNo")
                .Rows = .Rows + 1
                PR_PRGRN.MoveNext
                If PR_PRGRN.EOF Then Exit Do
             Loop
            
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalGRN
    Else
        Call SetErr("GRN Transactions not found.", vbCritical)
        txtLocCode.SetFocus
    End If
End Sub

Private Sub grdGRN_DblClick()
    With GrdGRN
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
        
        TxtItemCode = .TextMatrix(.Row, 1)
        txtqty = .TextMatrix(.Row, 2)
        txtUnitPrice = Val(.TextMatrix(.Row, 3))
        
        If MySeek(txtLocCode.Text & TxtItemCode.Text, "ItemCode", PR_PRItem) Then
            TxtFactor = PR_PRItem.Fields("ItemFactor")
            TxtUM = PR_PRItem.Fields("ItemUM")
        Else
            TxtFactor = 0
            TxtUM = 0
        End If
        PS_RowClicked = "Y"
    End With
End Sub

Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
    If LastKey(KeyCode) Then
       If txtvaluedate.Value < DateValue(Gs_Fnperiod) Or txtvaluedate.Value > DateValue(Gs_FnEndPeriod) Then
         Call SetErr("Invalid Period.", vbCritical)
         txtvaluedate.SetFocus
       Else
        txtLocCode.SetFocus
       End If
    End If
End Sub

Public Sub setfrmenv(ls_mode As String)
         txtLocCode.Enabled = IIf(ls_mode <> "D", True, False)
         TxtPartyCode.Enabled = IIf(ls_mode <> "D", True, False)
         TxtRemarks.Enabled = IIf(ls_mode <> "D", True, False)
         Frame2.Enabled = IIf(ls_mode <> "D", True, False)
End Sub
