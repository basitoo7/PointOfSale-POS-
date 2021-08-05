VERSION 5.00
Begin VB.Form frmAssetInfo 
   Caption         =   "Asset Information"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAssetInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   5265
      Begin VB.TextBox TxtLeaseRef 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   3
         Top             =   540
         Width           =   585
      End
      Begin VB.TextBox txtAssetDecs 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   5
         Top             =   900
         Width           =   3825
      End
      Begin VB.TextBox TxtAssetType 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1350
         MaxLength       =   2
         TabIndex        =   1
         Top             =   180
         Width           =   585
      End
      Begin VB.CommandButton Command11 
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
         Picture         =   "frmAssetInfo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2295
         MaxLength       =   35
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   2865
      End
      Begin VB.TextBox Txtchasis 
         Height          =   315
         Left            =   1350
         MaxLength       =   25
         TabIndex        =   10
         Top             =   2700
         Width           =   3825
      End
      Begin VB.TextBox TxtEngine 
         Height          =   315
         Left            =   1350
         MaxLength       =   25
         TabIndex        =   9
         Top             =   2340
         Width           =   3825
      End
      Begin VB.TextBox TxtReg 
         Height          =   315
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1980
         Width           =   3825
      End
      Begin VB.TextBox TxtSite 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1620
         Width           =   3825
      End
      Begin VB.TextBox TxtManu 
         Height          =   315
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1260
         Width           =   3825
      End
      Begin VB.CommandButton Command10 
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
         Picture         =   "frmAssetInfo.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox txtrecoverer 
         Height          =   315
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   12
         Top             =   3540
         Width           =   585
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
         Left            =   2340
         Picture         =   "frmAssetInfo.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3540
         Width           =   315
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2310
         MaxLength       =   35
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   2865
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Securitised By :"
         Height          =   210
         Index           =   7
         Left            =   150
         TabIndex        =   22
         Top             =   570
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Chasis # :"
         Height          =   210
         Index           =   6
         Left            =   570
         TabIndex        =   20
         Top             =   2730
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Engine # :"
         Height          =   210
         Index           =   5
         Left            =   600
         TabIndex        =   19
         Top             =   2370
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registration # :"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   2010
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Installation Site :"
         Height          =   210
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer :"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1290
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Index           =   1
         Left            =   390
         TabIndex        =   15
         Top             =   930
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Asset Type :"
         Height          =   210
         Index           =   0
         Left            =   375
         TabIndex        =   14
         Top             =   210
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmAssetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PO_CODE As Object
Public PO_DESC As Object

Private Sub Command10_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtAssetType
    Set PO_DESC = Text10
    
    GoTop frmLeaseAgree.PR_AssetType
    MyLookup.Caption = "Asset Types"
    MyLookup.FillGrid frmLeaseAgree.PR_AssetType, "AssetCode", "AssetName", 3
    MyLookup.Show 1
    
    If Trim(TxtAssetType) <> "" Then txtAssetType_KeyDown vbKeyReturn, vbKeyShift
    
End Sub
Private Sub command11_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtLeaseRef
    Set PO_DESC = Text11
    
    GoTop frmLeaseAgree.PR_leaseRef
    MyLookup.Caption = "Securitised By"
    MyLookup.FillGrid frmLeaseAgree.PR_leaseRef, "SecrtyCode", "SecrtyName", 3
    MyLookup.Show 1
    
    If Trim(TxtLeaseRef) <> "" Then txtLeaseRef_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_Activate()
    TxtAssetType.SetFocus
'  If frmLeaseAgree.Mode = "E" Then
'     TxtAssetType = AssetInfo(1)
'     TxtLeaseRef = AssetInfo(2)
'     txtAssetDecs = AssetInfo(3)
'     TxtSite = AssetInfo(4)
'     TxtManu = AssetInfo(5)
'     TxtReg = AssetInfo(6)
'     TxtEngine = AssetInfo(7)
'     Txtchasis = AssetInfo(8)
'  End If
End Sub

Private Sub txtAssetDecs_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then
      'AssetInfo(3) = txtAssetDecs.Text
      TxtManu.SetFocus
   ElseIf KeyCode = vbKeyPageUp Then
      TxtLeaseRef.SetFocus
   End If
End Sub

Private Sub txtAssetType_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyReturn And TxtAssetType.Text <> "" Then
      TxtAssetType.Text = DoPad(TxtAssetType.Text, TxtAssetType.MaxLength)
        If Not MySeek(TxtAssetType.Text, "AssetCODE", frmLeaseAgree.PR_AssetType) Then
           Call SetErr(Gs_RecNFMsg, vbCritical)
           TxtAssetType.SetFocus
        Else
           Text10.Text = frmLeaseAgree.PR_AssetType("AssetName")
           TxtReg.Enabled = IIf(frmLeaseAgree.PR_AssetType("AssetClass") = "V", True, False)
           TxtEngine.Enabled = IIf(frmLeaseAgree.PR_AssetType("AssetClass") = "V", True, False)
           Txtchasis.Enabled = IIf(frmLeaseAgree.PR_AssetType("AssetClass") = "V", True, False)
           frmAssetInfo.TxtLeaseRef.SetFocus
        End If
    ElseIf KeyCode = vbKeyF12 Then
       Call Command10_Click
    ElseIf KeyCode = vbKeyPageUp Then
       Unload Me
    End If
End Sub

Private Sub Txtchasis_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
      'AssetInfo(8) = Txtchasis.Text
      Me.Hide
  ElseIf KeyCode = vbKeyPageUp Then
    TxtEngine.SetFocus
  End If
End Sub

Private Sub TxtEngine_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
    'AssetInfo(7) = TxtEngine.Text
    Txtchasis.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
    TxtReg.SetFocus
  End If
End Sub

Private Sub txtLeaseRef_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyReturn And TxtLeaseRef.Text <> "" Then
      TxtLeaseRef.Text = DoPad(TxtLeaseRef.Text, TxtLeaseRef.MaxLength)
      
        If Not MySeek(TxtLeaseRef.Text, "SecrtyCode", frmLeaseAgree.PR_leaseRef) Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             TxtLeaseRef.SetFocus
        Else
             Text11.Text = frmLeaseAgree.PR_leaseRef("SecrtyName")
             txtAssetDecs.SetFocus
        End If
    ElseIf KeyCode = vbKeyF12 Then
       Call command11_Click
    ElseIf KeyCode = vbKeyPageUp Then
       TxtAssetType.SetFocus
    End If
End Sub

Private Sub TxtManu_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then
      'AssetInfo(4) = TxtManu.Text
      TxtSite.SetFocus
   ElseIf KeyCode = vbKeyPageUp Then
      txtAssetDecs.SetFocus
   End If
End Sub

Private Sub TxtReg_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     'AssetInfo(6) = TxtReg.Text
     TxtEngine.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     TxtSite.SetFocus
  End If
End Sub

Private Sub TxtSite_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then
      If TxtReg.Enabled Then
         'AssetInfo(5) = TxtSite.Text
         TxtReg.SetFocus
      Else
        Me.Hide
     End If
   ElseIf KeyCode = vbKeyPageUp Then
      TxtManu.SetFocus
   End If
End Sub
