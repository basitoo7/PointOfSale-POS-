VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmActiveUser 
   Caption         =   "Active User"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmActiveUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdLogOff 
      Caption         =   "&Log Off"
      Height          =   345
      Left            =   4440
      TabIndex        =   5
      Top             =   2070
      Width           =   885
   End
   Begin VB.Frame Frame2 
      Caption         =   "Active User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   6435
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   2040
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
      Begin VB.CommandButton Command2 
         Caption         =   "&Refresh"
         Height          =   345
         Left            =   3480
         TabIndex        =   4
         Top             =   2085
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   345
         Left            =   5400
         TabIndex        =   3
         Top             =   2085
         Width           =   885
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1770
         Left            =   105
         TabIndex        =   2
         Top             =   255
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3122
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "UserId"
            Caption         =   "User Name"
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
            DataField       =   "compname"
            Caption         =   "Computer Name"
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
            DataField       =   "logintime"
            Caption         =   "LoginTime"
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
            BeginProperty Column00 
               ColumnWidth     =   2039.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2340.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1560.189
            EndProperty
         EndProperty
      End
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
      TabIndex        =   1
      Top             =   2895
      Width           =   45
   End
End
Attribute VB_Name = "frmActiveUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MemUser As String

Private Sub CmdLogOff_Click()
 If MemUser <> "" Then
 gc_dbcon.Execute "Update SyUsers  set activestatus = 0,compname = '" & Trim(ls_CompName1) & "', logintime = '" & Time & "',logouttime = 'Still Active' where userid = '" & Trim(MemUser) & "'"
  Else
  MsgBox ("Please Select the User ...")
  End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Adodc1.ConnectionString = gc_dbcon
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "Select * from syusers where activestatus >0"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub DataGrid1_Click()

  With DataGrid1
      MemUser = .Text
     
 End With

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = gc_dbcon
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "Select * from syusers where ActiveStatus >0"

Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub


