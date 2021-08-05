VERSION 5.00
Begin VB.Form FrmFnYear 
   Caption         =   "Financial Years"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFnYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
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
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CheckBox chkActive 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   3240
         Picture         =   "FrmFnYear.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox TxtToDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtFromDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Active Year :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Year :"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmFnYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lb_BlnkMast As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Syfins As New Recordset
Dim PR_ChangeFY As New Connection
Dim grp_rs As New Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtFromDate
    Set PO_DESC = TxtToDate
    GoTop PR_Syfins
    MyLookup.Caption = "Financial Years"
    MyLookup.FillGrid PR_Syfins, "ffromdate", "ftodate", 14
    MyLookup.Show 1
    chkActive.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo localerr
'If MySeek(Gs_Fnperiod, "FFromDate", PR_Syfins) Then
   'PR_Syfins.Fields("factiveyear") = 0
   'PR_Syfins.Update
   If MySeek(TxtFromDate.Text, "FFromDate", PR_Syfins) Then
        'PR_Syfins.Fields("factiveyear") = 1
        'PR_Syfins.Update
        Gs_Fnperiod = PR_Syfins.Fields("ffromdate")
        Gs_FnEndPeriod = PR_Syfins.Fields("ftodate")
        Gs_DBName = PR_Syfins.Fields("dbname")
           
        Gs_msprovider = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & Gs_DBName & ";Data Source =" & Gs_DBDataSource
        If gc_dbcon.State = 1 Then gc_dbcon.Close
        gc_dbcon.Open Gs_msprovider
        gc_dbcon.CommandTimeout = 300
        Call changeODBCPath(Gs_DBName)
        
         
    Para_Rs.Open "Select * from SysComp", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    grp_rs.Open "Select * from Sysregs", gc_dbcon, adOpenStatic, adLockReadOnly
    GR_SMGroups.Open "Select * from Sys_UserGroups", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    GR_SMURights.Open "Select * from Sys_UserRights", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    
    'temp
   'Call chkdup1
  ' Delete Existed Temp Files
    PR_TempTables.Open "Select Name,Id,CrDate From SysObjects Where LEFT(name,4) = 'Tmp_' or name = 'coa'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    If Not grp_rs.EOF Then
       Gs_RegisterTo = grp_rs.Fields("GroupName")
    Else
        frmRegister.Show
        grp_rs.Requery
        If Not grp_rs.EOF Then
           Gs_RegisterTo = grp_rs.Fields("GroupName")
        End If
    End If
        
        Unload Me
   Else
        Call SetErr(Gs_RecNFMsg, vbCritical)
   End If
'Else
' Call SetErr("Current Active Year Not Found.", vbCritical)
'End If
Exit Sub

localerr:
Call MsgBox(Err.Description)
End Sub

Private Sub Form_Load()
Gs_msprovider = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=ecounts;Data Source =" & Gs_DBDataSource
If PR_ChangeFY.State = 1 Then PR_ChangeFY.Close
PR_ChangeFY.Open Gs_msprovider
  
  PR_Syfins.Open "Select * from SysFins Where compcode = '" & Gs_compcode & "' order by CompCode", PR_ChangeFY, adOpenDynamic, adLockOptimistic, 1
  
  If PR_Syfins.EOF Then
     Call SetErr("Data Not Found.", vbCritical)
     Unload Me
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Syfins.Close
End Sub
