VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManufacturer 
   Caption         =   "Manufacturer Setup"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManufacturer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7485
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
      Height          =   3285
      Left            =   45
      TabIndex        =   7
      Top             =   585
      Width           =   7380
      Begin VB.TextBox txtemail 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2925
         Width           =   5940
      End
      Begin VB.TextBox txtMCode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1305
         MaxLength       =   3
         TabIndex        =   26
         Tag             =   "SKIP"
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtContact 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2535
         Width           =   5940
      End
      Begin VB.TextBox txtcountryName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         Tag             =   "SKIP"
         Top             =   1770
         Width           =   5100
      End
      Begin VB.TextBox txtcountry 
         Height          =   330
         Left            =   1305
         MaxLength       =   3
         TabIndex        =   21
         Tag             =   "SKIP"
         Top             =   1770
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
         Left            =   1815
         Picture         =   "frmManufacturer.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "SKIP"
         Top             =   1770
         Width           =   315
      End
      Begin VB.TextBox txtcityname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "SKIP"
         Top             =   2145
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
         Left            =   1815
         Picture         =   "frmManufacturer.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "SKIP"
         Top             =   2145
         Width           =   315
      End
      Begin VB.TextBox txtCity 
         Height          =   330
         Left            =   1305
         MaxLength       =   3
         TabIndex        =   16
         Tag             =   "SKIP"
         Top             =   2145
         Width           =   495
      End
      Begin VB.TextBox txtmobile 
         Height          =   315
         Left            =   1305
         MaxLength       =   25
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2415
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   4830
         MaxLength       =   25
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2415
      End
      Begin VB.TextBox txtphoneoffice 
         Height          =   315
         Left            =   4815
         MaxLength       =   25
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1020
         Width           =   2430
      End
      Begin VB.TextBox txtnic 
         Height          =   315
         Left            =   1305
         MaxLength       =   25
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1005
         Width           =   2400
      End
      Begin VB.TextBox txtaddress 
         Height          =   315
         Left            =   1305
         MaxLength       =   200
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   615
         Width           =   5940
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "SKIP"
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
         Left            =   2085
         Picture         =   "frmManufacturer.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email :"
         Height          =   210
         Left            =   840
         TabIndex        =   28
         Top             =   2970
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contact Person :"
         Height          =   210
         Left            =   75
         TabIndex        =   24
         Top             =   2580
         Width           =   1200
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Country :"
         Height          =   210
         Left            =   615
         TabIndex        =   22
         Top             =   1815
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "City :"
         Height          =   210
         Left            =   885
         TabIndex        =   19
         Top             =   2190
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mobile :"
         Height          =   210
         Index           =   6
         Left            =   720
         TabIndex        =   15
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fax :"
         Height          =   210
         Index           =   5
         Left            =   4425
         TabIndex        =   14
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Phone :"
         Height          =   210
         Index           =   4
         Left            =   4260
         TabIndex        =   13
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NIC# :"
         Height          =   210
         Index           =   3
         Left            =   840
         TabIndex        =   12
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         Height          =   210
         Index           =   2
         Left            =   540
         TabIndex        =   11
         Top             =   645
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   210
         Index           =   0
         Left            =   2460
         TabIndex        =   10
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   795
         TabIndex        =   8
         Top             =   270
         Width           =   465
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
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
               Picture         =   "frmManufacturer.frx":0760
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManufacturer.frx":0BB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManufacturer.frx":1008
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManufacturer.frx":145C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManufacturer.frx":18B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManufacturer.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmManufacturer.frx":2458
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmManufacturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkSupp As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Supplier As New Recordset
Dim PR_Cities As New Recordset
Dim PR_Country As New Recordset
Dim PR_Dumy As New Recordset
Dim lb_found As Boolean
Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtmcode
    Set PO_DESC = txtdesc
    Gs_SQL = "Select MCode, Description from IC_Manufacturer"
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Manufacturer"
    MyLookupOLDB.Show 1
    
    If txtmcode.Text <> "" Then Call txtmcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

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



Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtCity <> "" Then
      txtCity.Text = DoPad(txtCity.Text, txtCity.MaxLength)
      PR_Cities.Open "Select * from cities where citycode = '" & txtCity & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1

      If Not PR_Cities.EOF Then
           txtcityname = PR_Cities("Cityname")
           If txtContact.Enabled Then txtContact.SetFocus
        Else
           Call SetErr(Gs_RecNFMsg, vbCritical)
           txtCity.SetFocus
      End If
     PR_Cities.Close
End If
End Sub

Private Sub txtContact_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtemail.SetFocus
End Sub

Private Sub txtcountry_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtcountry <> "" Then
      txtcountry.Text = DoPad(txtcountry.Text, txtcountry.MaxLength)
      PR_Country.Open "Select * from Country where Countrycode = '" & txtcountry & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1

      If Not PR_Country.EOF Then
           txtcountryName = PR_Country("Countryname")
           txtCity.SetFocus
        Else
           Call SetErr(Gs_RecNFMsg, vbCritical)
           txtcountry.SetFocus
      End If
     PR_Country.Close
End If
End Sub



Private Sub txtstreg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtopeningbal.SetFocus
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Supplier, frmSupplier, txtmcode, txtdesc, Para_Rs, "IC_PartyCnt", 6, "SupplierCode", "Description", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("GLSUB00001")
  SetToolBar(2) = chkRights("GLSUB00002")
  SetToolBar(3) = chkRights("GLSUB00003")
  SetToolBar(4) = chkRights("GLSUB00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
 End Sub
Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtaddress.SetFocus
End Sub
Private Sub txtaddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtnic.SetFocus
End Sub
Private Sub txtNIC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtphoneoffice.SetFocus
End Sub
Private Sub txtphoneoffice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtmobile.SetFocus
End Sub
Private Sub txtfax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtcountry.SetFocus
End Sub
Private Sub txtmobile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtFax.SetFocus
End Sub


Private Sub txtmcode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Trim(txtmcode.Text) <> "" Then
         txtmcode.Text = DoPad(txtmcode.Text, txtmcode.MaxLength)
         PR_Supplier.Open "Select * from IC_Manufacturer where compcode = '" & Gs_compcode & "'  and mcode = '" & txtmcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1

       Select Case Mode
            Case "A"
                If Not PR_Supplier.EOF Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   txtmcode.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If PR_Supplier.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Call ClearVal
                   txtmcode.SetFocus
                Else
                   Call SetVal
                End If
            End Select
          PR_Supplier.Close
  ElseIf KeyCode = vbKeyReturn And Trim(txtmcode.Text) = "" Then
    txtmcode = ""
    txtdesc = ""
  End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
    Else
      cmdLookup.Enabled = True
    End If
    Mode = DentMode(Mode, Button.Index, PR_Supplier, frmManufacturer, txtmcode, txtdesc, Para_Rs, "IC_PartyCnt", 6, "SupplierCode", "Description", 1, False, Toolbar1)
    If Mode = "A" Then
       txtmcode = maxtranscode
       txtdesc.SetFocus
    End If

End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim ls_sql As String
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
            ls_sql = "INSERT into IC_Manufacturer(Compcode, MCode, Description, Address,NicNo, CountryCode, CityCode, Phoneoffice, Fax, CP, Email,Mobile)"
            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & txtmcode.Text & "','" & txtdesc.Text & "','" & txtaddress.Text & "','" & txtnic.Text & "','" & txtcountry.Text & "','" & txtCity.Text & "','" & txtphoneoffice.Text & "','" & txtFax.Text & "','" & txtContact.Text & "','" & txtemail.Text & "','" & txtmobile.Text & "')"
            gc_dbcon.Execute ls_sql
           Case "E"
             ls_sql = " UPDATE IC_Manufacturer SET Description= '" & txtdesc.Text & "',Address= '" & txtaddress.Text & "',NicNo= '" & txtnic.Text & "',Phoneoffice= '" & txtphoneoffice.Text & "',Fax= '" & txtFax.Text & "',Mobile= '" & txtmobile.Text & "', countryCode  = '" & txtcountry & "', CityCode  = '" & txtCity & "', CP  = '" & txtContact & "', email  = '" & Trim(txtemail) & "' WHERE  MCode= '" & txtmcode.Text & "' and Compcode = '" & Gs_compcode & "'"
             gc_dbcon.Execute ls_sql
           Case "D"
              gc_dbcon.Execute "DELETE FROM IC_Manufacturer WHERE  MCode= '" & txtmcode.Text & "' and Compcode = '" & Gs_compcode & "'"
           
     End Select
gc_dbcon.CommitTrans

Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Function maxtranscode() As String
PR_Dumy.Open "select max(MCode) as transcode from IC_Manufacturer where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & PR_Dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
PR_Dumy.Close
End Function
Public Sub ClearVal()
     txtmcode = ""
     txtdesc = ""
     txtCrValue = ""
     txtCrQty = ""
     'Option1(0).Value = True
End Sub

Private Sub SetVal()
     txtdesc = Trim(PR_Supplier("Description") & "")
     txtaddress = Trim(PR_Supplier("Address") & "")
     txtnic = Trim(PR_Supplier("NicNo") & "")
     txtphoneoffice = Trim(PR_Supplier("Phoneoffice") & "")
     txtFax = Trim(PR_Supplier("Fax") & "")
     txtmobile = Trim(PR_Supplier("Mobile") & "")
     txtContact = Trim(PR_Supplier("CP") & "")
     txtemail = Trim(PR_Supplier("Email") & "")
     txtcountry = PR_Supplier("CountryCode") & ""
     If txtcountry <> "" Then Call txtcountry_KeyDown(vbKeyReturn, vbKeyShift)
     txtCity = PR_Supplier("CityCode") & ""
     If txtCity <> "" Then Call txtCity_KeyDown(vbKeyReturn, vbKeyShift)
     
End Sub
Public Function ChkInputs() As Boolean
    If Trim(txtmcode.Text) = "" Then
       Call MsgBox("Enter Code!!!", vbCritical)
       txtmcode.SetFocus
       ChkInputs = False
    ElseIf Trim(txtdesc.Text) = "" Then
       Call MsgBox("Enter Description!!!", vbCritical)
       txtdesc.SetFocus
       ChkInputs = False
    ElseIf Trim(txtnic.Text) = "" Then
       Call MsgBox("Enter NIC Number!!!", vbCritical)
       txtnic.SetFocus
       ChkInputs = False
    ElseIf txtphoneoffice.Text = "" Then
       Call MsgBox("Enter Phone Number!!!", vbCritical)
       txtphoneoffice.SetFocus
       ChkInputs = False
    ElseIf txtFax.Text = "" Then
       Call MsgBox("Enter Fax Number!!!", vbCritical)
       txtFax.SetFocus
       ChkInputs = False
    ElseIf txtmobile.Text = "" Then
       Call MsgBox("Enter Mobile Number!!!", vbCritical)
       txtmobile.SetFocus
       ChkInputs = False
    ElseIf txtContact.Text = "" Then
       Call MsgBox("Enter Contact Person!!!", vbCritical)
       txtContact.SetFocus
       ChkInputs = False
    ElseIf txtemail.Text = "" Then
       Call MsgBox("Enter Email Address!!!", vbCritical)
       txtemail.SetFocus
       ChkInputs = False
    ElseIf txtcountry.Text = "" Then
       Call MsgBox("Enter/Select Country Code!!!", vbCritical)
       txtcountry.SetFocus
       ChkInputs = False
    ElseIf txtCity.Text = "" Then
       Call MsgBox("Enter/Select City Code!!!", vbCritical)
       txtCity.SetFocus
       ChkInputs = False
   Else
       ChkInputs = True
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtdesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub
