VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MyLookupSaleOLDB 
   Caption         =   "Look up :"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MyLookupSaleOLDB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdtree 
      Caption         =   "&Tree"
      Height          =   300
      Left            =   10800
      TabIndex        =   5
      Top             =   5805
      Visible         =   0   'False
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
      Height          =   5025
      Left            =   30
      TabIndex        =   3
      Top             =   840
      Width           =   11580
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4620
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   8149
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483645
         HeadLines       =   1
         RowHeight       =   23
         TabAction       =   1
         WrapCellPointer =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
   End
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
      Height          =   735
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   9840
      Begin VB.TextBox SeekText 
         Height          =   420
         Left            =   720
         MaxLength       =   50
         TabIndex        =   0
         Top             =   195
         Width           =   8970
      End
      Begin VB.Label Label1 
         Caption         =   "&Text :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9945
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblmsg 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   45
      TabIndex        =   4
      Top             =   2895
      Width           =   45
   End
End
Attribute VB_Name = "MyLookupSaleOLDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ln_res As Integer
Dim PR_Sub1 As New Recordset
Public ls_Accountno As String

Private Sub cmdtree_Click()
On Error GoTo LocalErr
    ls_Accountno = DataGrid1.Columns(0).Text
    Mytree.lbldetail = DataGrid1.Columns(1).Text
    Mytree.Show 1
    Exit Sub
LocalErr:
End Sub

Private Sub DataGrid1_Click()
'If Gs_Subon Then Call DataGrid1_KeyDown(vbKeyUp, vbKeyShift)
If Gs_Subon Then Call DataGrid1_KeyUp(vbKeyUp, vbKeyShift)
DataGrid1.Refresh
End Sub

Private Sub DataGrid1_GotFocus()
'If Gs_Subon Then Call DataGrid1_KeyDown(vbKeyUp, vbKeyShift)
If Gs_Subon Then Call DataGrid1_KeyUp(vbKeyUp, vbKeyShift)
DataGrid1.Refresh
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
Dim sortField As String
Dim sortString As String

sortField = DataGrid1.Columns(ColIndex).Caption
If InStr(Adodc1.Recordset.Sort, "Asc") Then
    sortString = sortField & " Desc"
Else
    sortString = sortField & " Asc"
End If
Adodc1.Recordset.Sort = sortString
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LocalErr

If KeyCode = vbKeyReturn Then
    If Adodc1.Recordset.RecordCount >= 1 Then
        PO_AnyForm.PO_CODE = Trim(DataGrid1.Columns(0).Text)
        PO_AnyForm.PO_DESC = Trim(DataGrid1.Columns(1).Text)
        Unload Me
        
    End If
ElseIf KeyCode = vbKeyPageUp Then
    SeekText.SetFocus
'ElseIf (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) And Gs_Subon Then
'    If MySeek(Trim(Left(DataGrid1.Columns(0).Text, 10)), "Findfld", PR_Sub1) Then
'        lblmsg.Caption = PR_Sub1("Acct_desc")
'        DataGrid1.ToolTipText = PR_Sub1("Acct_desc")
'    End If
End If

Exit Sub
LocalErr:
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LocalErr

'If KeyCode = vbKeyReturn Then
'    If Adodc1.Recordset.RecordCount >= 1 Then
'        PO_AnyForm.PO_CODE = DataGrid1.Columns(0).Text
'        PO_AnyForm.PO_DESC = DataGrid1.Columns(1).Text
'        Unload Me
'    End If
'ElseIf KeyCode = vbKeyPageUp Then
'    SeekText.SetFocus

If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) And Gs_Subon Then
    If MySeek(Trim(Left(DataGrid1.Columns(0).Text, gn_sublen(0) + gn_sublen(1) + gn_sublen(2) + gn_sublen(3))), "Findfld", PR_Sub1) Then
        lblmsg.Caption = PR_Sub1("Acct_desc")
        DataGrid1.ToolTipText = PR_Sub1("Acct_desc")
      
    End If
End If

Exit Sub
LocalErr:
ln_res = SetErr("Critical error occurred Please report to MIS Department", vbCritical)

End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    If Adodc1.Recordset.RecordCount > 1 Then DataGrid1_KeyDown vbKeyReturn, vbKeyShift
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
    
End If
End Sub

Private Sub Form_Load()
On Error GoTo LocalErr
Adodc1.ConnectionString = gc_dbcon
Adodc1.CommandType = adCmdText

Adodc1.RecordSource = Gs_SQL + Gs_OtherPara + Gs_ExtraPara + " " + Gs_OrderBy

Set DataGrid1.DataSource = Adodc1
DataGrid1.Columns(1).Width = 6000
DataGrid1.Columns(2).Width = 900
DataGrid1.Columns(3).Width = 800
DataGrid1.Columns(4).Width = 900
 
 DataGrid1.Refresh
 
 'If Gs_Subon = True Then
 '   PR_Sub1.Open "select gl_sub2.*,ltrim(rtrim(gl_sub2.acct_sub1))+ltrim(rtrim(gl_sub2.acct_sub2)) as Findfld from gl_sub2 where compcode = '" & Gs_compcode & "' order by  Findfld ", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
 '   cmdtree.Visible = True
 'End If
 
  Exit Sub
LocalErr:
ln_res = SetErr("Critical error occurred Please report to MIS Department", vbCritical)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Gs_Subon = True Then
         PR_Sub1.Close
    End If
    Gs_SQL = ""
    Gs_Subon = False
    Gs_FindFld = ""
    Gs_OrderBy = ""
    Gs_OtherPara = ""
    
End Sub

Private Sub SeekText_Change()
On Error GoTo LocalErr
If SeekText.Text <> "" Then
   
  If IsNumeric(SeekText.Text) Then
   Gs_FindFld = "CUSTOMCODE"
   Adodc1.RecordSource = Gs_SQL & IIf(Len(Gs_OtherPara) > 0, Gs_OtherPara & " And " & UCase(Gs_FindFld) & " like '" & UCase(SeekText) & "%" & "'" + " " + Gs_OrderBy, " Where " & UCase(Gs_FindFld) & " like '" & UCase(SeekText) & "%" & "'" + " " + Gs_OrderBy)
  Else
   Gs_FindFld = "Description"
    Adodc1.RecordSource = Gs_SQL & IIf(Len(Gs_OtherPara) > 0, Gs_OtherPara & " And " & UCase(Gs_FindFld) & " like '" & UCase(SeekText) & "%" & "'" + " " + Gs_OrderBy, " Where " & UCase(Gs_FindFld) & " like '" & UCase(SeekText) & "%" & "'" + " " + Gs_OrderBy)
  End If
   Adodc1.Refresh
Else
     Adodc1.RecordSource = Gs_SQL + Gs_OtherPara + " " + Gs_OrderBy
   ' Adodc1.Refresh
End If

Set DataGrid1.DataSource = Adodc1
DataGrid1.Columns(1).Width = 6000
DataGrid1.Columns(2).Width = 900
DataGrid1.Columns(3).Width = 800
DataGrid1.Columns(4).Width = 900

 DataGrid1.SetFocus
 DataGrid1.Refresh
 SeekText.SetFocus
Exit Sub
LocalErr:
ln_res = SetErr("Critical error occurred Please report to MIS Department", vbCritical)
End Sub

Private Sub SeekText_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
       DataGrid1.SetFocus
End If
End Sub
