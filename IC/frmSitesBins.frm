VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSitesBins 
   Caption         =   "Sites Bins"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSitesBins.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5010
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
      Height          =   2805
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   4935
      Begin VB.TextBox txtsitedesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   210
         Width           =   2610
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFFF00&
         Height          =   330
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "SKIPN"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   5
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
         Left            =   1875
         Picture         =   "frmSitesBins.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid GrdGRN 
         Height          =   1710
         Left            =   45
         TabIndex        =   7
         Top             =   1035
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3016
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bin Description :"
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Site Code :"
         Height          =   210
         Left            =   405
         TabIndex        =   2
         Top             =   240
         Width           =   780
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5010
      _ExtentX        =   8837
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
               Picture         =   "frmSitesBins.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSitesBins.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSitesBins.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSitesBins.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSitesBins.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSitesBins.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSitesBins.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSitesBins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Sites As Recordset

Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String


Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocation
    Set PO_DESC = TxtsiteDesc
    Gs_SQL = "Select SiteCode, Description from IC_Sites "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Company Sites/Stores"
    MyLookupOLDB.Show 1
    If txtLocation <> "" Then Call TxtLocation_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SRSIB00001")
  SetToolBar(2) = chkRights("SRSIB00002")
  SetToolBar(3) = chkRights("SRSIB00003")
  SetToolBar(4) = chkRights("SRSIB00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  
  Set PR_Sites = New Recordset
   
  PR_Sites.Open "Select * from IC_Sites where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
   
  PB_BlnkLoca = IIf(PR_Sites.EOF, True, False)
  
  InitializeGrid

End Sub


Private Sub Form_Unload(Cancel As Integer)
    PR_Sites.Close
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call GrdGRN_DblClick
   If KeyCode = vbKeyDelete Then
       With GrdGRN
          If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
             .RemoveItem .Row
             If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
             End If
       End With
   End If
End Sub


Private Sub GrdGRN_DblClick()
  With GrdGRN
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
       txtdesc = .TextMatrix(.Row, 1)
       PS_RowClicked = "Y"
       txtdesc.SetFocus
    End With
End Sub


Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(txtdesc) <> "" Then
    Call AddToGrid
ElseIf KeyCode = vbKeyReturn And Trim(txtdesc) = "" Then
    Call MsgBox("Enter Bin Description !!!", vbCritical)
    txtdesc.SetFocus
End If
End Sub

Private Sub TxtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn Then
         
         txtLocation.Text = IIf(IsNumeric(txtLocation.Text), DoPad(UCase(txtLocation.Text), txtLocation.MaxLength), UCase(txtLocation.Text))
         lb_found = MySeek(txtLocation.Text, "SiteCode", PR_Sites)
         Select Case Mode
         
         Case "A"
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                   txtLocation.SetFocus
                Else
                   TxtsiteDesc = PR_Sites("Description")
                   txtdesc.SetFocus
                End If
         Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                   txtLocation.SetFocus
                Else
                   TxtsiteDesc = PR_Sites("Description")
                   txtdesc.SetFocus
                   SetVal
                End If
         End Select
            
         
 End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Sites, Me, txtLocation, txtdesc, "X", "CompCount", 3, "SiteCode", "Description", 1, False, Toolbar1)
    End If
    If Button.Index = 6 Then InitializeGrid
End Sub
Private Sub AddToGrid()
Dim ln_cnt As Integer
            If Trim(txtdesc) <> "" Then
                    If PS_RowClicked = "" Then
                        If PI_SrNo = 0 Then
                            PI_SrNo = 1
                        Else
                            PI_SrNo = PI_SrNo + 1
                         End If
                     End If
        
                    If Trim(txtdesc) <> "" Then
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
                                   .TextMatrix(.Row, 1) = Trim(txtdesc)
                                                
                                txtdesc.Text = ""
                                
                                PS_RowClicked = ""
                        End With
                     End If
                        
                        txtdesc.SetFocus
                   
        Else
            Call SetErr("Enter bin Description", vbCritical)
            txtdesc.SetFocus
       End If
End Sub

Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Bin Description"
        .ColWidth(1) = 4000
        .Redraw = True
    End With
    PI_SrNo = 0
    PI_CurRow = 0
    PS_RowClicked = ""
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim ls_sql As String
PB_BlnkLoca = False
gc_dbcon.BeginTrans
     Select Case Mode
           Case "D"
              ls_sql = "DELETE FROM IC_SitesBins WHERE SiteCode = '" & txtLocation.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
            
           Case Else
              ls_sql = "DELETE FROM IC_SitesBins WHERE SiteCode = '" & txtLocation.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
            
              With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      ls_sql = "INSERT into IC_Sitesbins(compcode,SiteCode,BinCode,Description) VALUES ('" & Gs_compcode & "','" & txtLocation.Text & "'," & .TextMatrix(ln_cnt, 0) & " , '" & RepApp(.TextMatrix(ln_cnt, 1)) & "')"
                      gc_dbcon.Execute ls_sql
                    Next
              End With
              
              
     End Select

gc_dbcon.CommitTrans
PR_Sites.Requery
InitializeGrid
txtLocation.Text = ""
TxtsiteDesc = ""


Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Public Sub ClearVal()
     txtLocation = ""
     txtdesc = ""
End Sub

Private Sub SetVal()
On Error GoTo LocalErr

Dim pr_dumyloadtrans As New Recordset

InitializeGrid
    
    pr_dumyloadtrans.Open "select * from Ic_sitesBins where compcode = '" & Gs_compcode & "' and SiteCode = '" & txtLocation & "'", gc_dbcon, adOpenStatic, adLockReadOnly
   
    If Not pr_dumyloadtrans.EOF Then
        With GrdGRN
            Do While Not pr_dumyloadtrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(pr_dumyloadtrans("Description") & "")
                .Rows = .Rows + 1
                pr_dumyloadtrans.MoveNext
                If pr_dumyloadtrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
    End If
    pr_dumyloadtrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)
     
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtLocation.Text) = txtLocation.MaxLength And PI_SrNo > 0 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtdesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub
