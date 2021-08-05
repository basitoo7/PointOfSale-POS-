VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "frmItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1170
      Width           =   915
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   1170
      Width           =   915
   End
   Begin VB.TextBox txtLDesc 
      Height          =   315
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame7 
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3435
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2400
         Picture         =   "frmItems.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   315
      End
      Begin VB.TextBox txtIDesc 
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtItemClass 
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
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtLocCode 
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   2400
         Picture         =   "frmItems.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item Class :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   8
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Location Code :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1125
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1635
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Description"
            TextSave        =   "Description"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport rptItems 
      Left            =   60
      Top             =   1320
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
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkLoc As Boolean
Dim Pb_BlnIClass As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Locations As Recordset
Dim PR_ItemClass As Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
    With rptItems
        .ReportFileName = App.Path & Gs_ICRepoPath & "\Items.RPT"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .SelectionFormula = "{IC_Item.CompCode} = '" & Gs_compcode & "'"
        If Len(Trim(txtLocCode)) > 0 Then .SelectionFormula = .SelectionFormula & " AND {IC_Item.LocationCode} = '" & txtLocCode & "'"
        If Len(Trim(txtItemClass)) > 0 Then .SelectionFormula = .SelectionFormula & " AND {IC_Item.ItemClass} = '" & txtItemClass & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
Exit Sub
LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command3_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = txtLDesc
    
    GoTop PR_Locations
    MyLookup.Caption = "Locations"
    MyLookup.FillGrid PR_Locations, "LocationCode", "Description", 5
    MyLookup.Show 1
    
    If Len(txtLocCode) > 0 Then txtLocCode_KeyDown vbKeyReturn, vbKeyShift

End Sub
Private Sub Command1_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtItemClass
    Set PO_DESC = txtIDesc
    
    GoTop PR_Locations
    MyLookup.Caption = "Classes"
    MyLookup.FillGrid PR_ItemClass, "ItemClass", "Description", 5
    MyLookup.Show 1
    
    If Len(txtLocCode) > 0 Then txtItemClass_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_Load()
  
  Set PR_Locations = New Recordset
  Set PR_ItemClass = New Recordset
  
  PR_Locations.Open "Select * from IC_Locations where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_ItemClass.Open "Select * from IC_ItemClass where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  Pb_BlnkLoc = IIf(PR_Locations.EOF, True, False)
  Pb_BlnIClass = IIf(PR_ItemClass.EOF, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Locations.Close
    PR_ItemClass.Close
End Sub

Private Sub txtItemClass_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If Lastkey(KeyCode) And txtItemClass.Text <> "" Then
         txtLocCode = UCase(txtItemClass)
         lb_found = MySeek(txtItemClass.Text, "ItemClass", PR_ItemClass)
   
         If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtLocCode.SetFocus
         Else
            StatusBar1.Panels(2) = PR_ItemClass.Fields("Description")
            cmdGenerate.SetFocus
         End If
  ElseIf Lastkey(KeyCode) And txtItemClass.Text = "" Then
         cmdGenerate.SetFocus
  End If
End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If Lastkey(KeyCode) And txtLocCode.Text <> "" Then
         txtLocCode = UCase(txtLocCode)
         lb_found = MySeek(txtLocCode.Text, "LocationCode", PR_Locations)
   
         If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtLocCode.SetFocus
         Else
            StatusBar1.Panels(2) = PR_Locations.Fields("Description")
            cmdGenerate.SetFocus
         End If
  ElseIf Lastkey(KeyCode) And txtLocCode.Text = "" Then
         cmdGenerate.SetFocus
  End If
End Sub


