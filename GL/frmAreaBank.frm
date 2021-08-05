VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAreaBank 
   Caption         =   "Area Bank Setup"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4905
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4695
      Begin VB.TextBox m_Bank_Manager 
         Height          =   285
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Bank Manager Name:"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox m_Bank_Name 
         Height          =   285
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Bank Name :"
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox m_Bank_Code 
         BackColor       =   &H00FFFF00&
         CausesValidation=   0   'False
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "frmAreaBank.frx":0000
         Left            =   1560
         List            =   "frmAreaBank.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Bank Code :"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Manager Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Name :"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   " Bank Code :"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   1111
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
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
         Left            =   2040
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
               Picture         =   "frmAreaBank.frx":0004
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAreaBank.frx":0458
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAreaBank.frx":08AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAreaBank.frx":0D00
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAreaBank.frx":1154
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAreaBank.frx":15A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAreaBank.frx":1CFC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAreaBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BlankBank As Boolean
Dim Mode As String
Dim Bank_Rs As Recordset
Private Sub Form_Load()

  Top = (Screen.Height - Height) / 2
  Left = (Screen.Width - Width) / 2
  
  Toolbar1.Buttons(1).Enabled = mSetups(11)
  Toolbar1.Buttons(2).Enabled = mSetups(12)
  Toolbar1.Buttons(3).Enabled = mSetups(13)
  Toolbar1.Buttons(5).Enabled = mSetups(14)
  
  Set Para_Rs = New Recordset
  Set Bank_Rs = New Recordset
  
  Bank_Rs.Open "Select * from Area_Banks order by Bank_Code", CaneConn, adOpenDynamic, adLockOptimistic, 1
  Para_Rs.Open "Select Bank_Nos from ParaCount", CaneConn, adOpenDynamic, adLockOptimistic, 1
  
  BlankBank = IIf(Bank_Rs.EOF, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Bank_Rs.Close
    Para_Rs.Close
End Sub

Private Sub m_Bank_Code_GotFocus()
  If Mode = "A" Then
       m_Bank_Code.Clear
       m_Bank_Code = Format(LTrim(Str(Para_Rs("Bank_Nos") + 1)), "00")
       m_Bank_Code.Enabled = False
       m_Bank_Name.SetFocus
  End If
End Sub

Private Sub m_Bank_Code_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MbCode As String
    
 If Not Mode = "A" And KeyCode = vbKeyReturn Then
        MbCode = IIf(Len(LTrim(m_Bank_Code)) = 2, LTrim(m_Bank_Code), Left(m_Bank_Code, 2))
       
       If Not MySeek(MbCode, "Bank_Code", Bank_Rs) Then
          MsgBox "Required record does not exist.", vbCritical, "Visual Cane Says!"
          Cancel = True
       Else
          Call SetVal
          m_Bank_Name.Locked = IIf(Mode = "D", True, False)
          m_Bank_Manager.Locked = IIf(Mode = "D", True, False)
       End If
  End If
End Sub
Private Sub m_Bank_Name_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
         m_Bank_Manager.SetFocus
     End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If BlankBank And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical, "Visual Cane Says:"
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, Bank_Rs, frmAreaBank, m_Bank_Code, m_Bank_Name, "Bank_Nos", 2, "Bank_Code", "Bank_Name", 0, True, True)
    End If
End Sub

Public Sub SaveValues()
     Bank_Rs("Bank_Code") = IIf(Mode = "A", m_Bank_Code, Bank_Rs("Bank_Code"))
     Bank_Rs("Bank_Name") = m_Bank_Name
     Bank_Rs("Bank_Manager") = m_Bank_Manager
     BlankBank = False
End Sub
Public Sub ClearVal()
     m_Bank_Name = ""
     m_Bank_Manager = ""
End Sub

Private Sub SetVal()
   m_Bank_Name = Bank_Rs("Bank_Name")
   m_Bank_Manager = Bank_Rs("Bank_Manager")
   m_Bank_Name.SetFocus
End Sub


Public Function ChkInputs() As Boolean
    If Len(m_Bank_Name) > 0 Then
       ChkInputs = True
    Else
       MsgBox "Incomplete Data found:", vbCritical, "Visual Cane Says:"
       ChkInputs = False
    End If
End Function

