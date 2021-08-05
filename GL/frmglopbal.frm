VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmopbal 
   Caption         =   "Opening Balances."
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5310
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Tax Year :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
      Begin MSMask.MaskEdBox txttdr_amount 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txttcr_amount 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Credit Amount :"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Debit Amount :"
         Height          =   210
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Financial Year :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   2655
      Begin MSMask.MaskEdBox txtfdr_amount 
         Height          =   280
         Left            =   1320
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfcr_amount 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Credit Amount :"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Debit Amount :"
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   480
         Width           =   1050
      End
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
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   5295
      Begin VB.TextBox Txtdescrip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   600
         Width           =   3615
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
         Left            =   4800
         Picture         =   "frmglopbal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtAcctNo 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Enter Company code"
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   50
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account # :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   840
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5310
      _ExtentX        =   9366
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
               Picture         =   "frmglopbal.frx":0172
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmglopbal.frx":05C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmglopbal.frx":0A1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmglopbal.frx":0E6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmglopbal.frx":12C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmglopbal.frx":1716
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmglopbal.frx":1E6A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmopbal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PB_BlnkAct As Boolean
Dim Mode As String
Dim PR_GlActNo As Recordset

Private Sub Form_Load()
  
  Toolbar1.Buttons(1).Enabled = True
  Toolbar1.Buttons(2).Enabled = True
  Toolbar1.Buttons(3).Enabled = True
  Toolbar1.Buttons(5).Enabled = True
  
  Set PR_GlActNo = New Recordset
   
  PR_GlActNo.Open "Select * from Gl_Detail", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
   
  PB_BlnkAct = IIf(PR_GlActNo.EOF, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GlActNo.Close
End Sub

Private Sub txtAcctNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtAcctNo.Text) > 0 Then
          PR_GlActNo.Requery
          
         lb_found = MySeek(txtAcctNo.Text, "AccountNo", PR_GlActNo)
         txtfdr_amount.Enabled = IIf(Mode <> "D", True, False)
         txtfcr_amount.Enabled = IIf(Mode <> "D", True, False)
         txttdr_amount.Enabled = IIf(Mode <> "D", True, False)
         txttcr_amount.Enabled = IIf(Mode <> "D", True, False)
         
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                   Call ClearVal
                   txtAcctNo.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtAcctNo.Enabled = False
                      txtfdr_amount.SetFocus
                   End If
                End If
       End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_BlnkAct And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlActNo, frmopbal, txtAcctNo, txtfdr_amount, "X", "X", 3, "X", "X", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkComp = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A" Or "E"
              cntsql.CommandText = "UPDATE Gl_Detail SET FDr_Amount= '" & txtfdr_amount.Text & "',Fcr_amount = '" & txtfcr - amount & "',tdr_Amount= '" & txttdr_amount & "',tcr_amount = '" & txttcr_amount & "' WHERE  compcode = '" & Gs_compcode & "' and Acct_sub0= '" & Txtcctno.Text & "'"
              txtsub0.Enabled = True
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "UPDATE Gl_Detail SET FDr_Amount= '" & "0" & "',Fcr_amount = '" & "0" & "',tdr_Amount= '" & "0" & "',tcr_amount = '" & "0" & "' WHERE  compcode = '" & Gs_compcode & "' and Acct_sub0= '" & Txtcctno.Text & "'"
              cntsql.Execute
           
     End Select
End Sub
Public Sub ClearVal()
     txtAcctNo = ""
     Txtdescrip.Text = ""
     txtfdr_amount = ""
     txtfcr_amount = ""
     txttdr_amount = ""
     txttcr_amount = ""
End Sub

Private Sub SetVal()
     Txtdescrip.Text = PR_GlActNo("Acct_Desc")
     txtfdr_amount = PR_GlActNo("fdr_Amount")
     txtfcr_amount = PR_GlActNo("fCr_Amount")
     txttdr_amount = PR_GlActNo("Tdr_Amount")
     txttcr_amount = PR_GlActNo("TCr_Amount")
     
End Sub
Public Function ChkInputs() As Boolean
    If Val(txtAcctNo.Text) > 0 And (Val(txtfdr_amount.Text) > 0 Or Val(txtfcr_amount.Text) > 0 Or Val(txttdr_amount.Text) > 0 Or Val(txttcr_amount.Text) > 0) Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function



