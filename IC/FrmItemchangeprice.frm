VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmItemchangeprice 
   Caption         =   "Item Price Schedule"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmItemchangeprice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5040
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
      Height          =   2925
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   5025
      Begin VB.TextBox txtmcode 
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   23
         Top             =   975
         Width           =   435
      End
      Begin VB.TextBox txtmdesc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         MaxLength       =   6
         TabIndex        =   22
         Top             =   975
         Width           =   2715
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
         Left            =   1770
         Picture         =   "FrmItemchangeprice.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   975
         Width           =   315
      End
      Begin VB.TextBox txttotalsaleprice 
         Height          =   315
         Left            =   3750
         MaxLength       =   25
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1185
      End
      Begin VB.TextBox txtsaleprice 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2115
         Width           =   1110
      End
      Begin VB.TextBox txttotalpurchaseprice 
         Height          =   315
         Left            =   3750
         MaxLength       =   25
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1725
         Width           =   1170
      End
      Begin VB.TextBox txtpurchaseprice 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1725
         Width           =   1110
      End
      Begin VB.TextBox txtnoofitem 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1110
      End
      Begin VB.TextBox txtweight 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2505
         Width           =   1110
      End
      Begin VB.TextBox Textx 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   4620
         MaxLength       =   35
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   4
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
         Left            =   1815
         Picture         =   "FrmItemchangeprice.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtcode 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   3
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
      Begin MSComCtl2.DTPicker Dtpopeningdate 
         Height          =   300
         Left            =   3765
         TabIndex        =   24
         Top             =   2535
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Format          =   65667073
         CurrentDate     =   39058
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Date :"
         Height          =   195
         Left            =   2595
         TabIndex        =   25
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit Price (Sale) :"
         Height          =   450
         Left            =   2730
         TabIndex        =   20
         Top             =   2055
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit Price (Purchase) :"
         Height          =   450
         Left            =   2730
         TabIndex        =   18
         Top             =   1650
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Price Per Items (Sale) :"
         Height          =   450
         Left            =   15
         TabIndex        =   13
         Top             =   2070
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Price Per Items (Purchase) :"
         Height          =   450
         Left            =   15
         TabIndex        =   12
         Top             =   1650
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No of Items :"
         Height          =   210
         Left            =   375
         TabIndex        =   11
         Top             =   1335
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unit :"
         Height          =   210
         Left            =   930
         TabIndex        =   10
         Top             =   990
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Weight :"
         Height          =   210
         Left            =   705
         TabIndex        =   9
         Top             =   2550
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   180
         TabIndex        =   7
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   375
         TabIndex        =   6
         Top             =   600
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
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
            Caption         =   "Refresh"
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
               Picture         =   "FrmItemchangeprice.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemchangeprice.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemchangeprice.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemchangeprice.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemchangeprice.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemchangeprice.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemchangeprice.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmItemchangeprice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_itemsetup As New Recordset
Dim PR_itemmsetup As New Recordset
Dim pr_dumy As New Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCode
    Set PO_DESC = txtDesc
    
    GoTop PR_itemsetup
    MyLookup.Caption = "Item Setup"
    MyLookup.FillGrid PR_itemsetup, "itemcode", "description", txtCode.MaxLength
    MyLookup.Show 1
    
    If Len(txtCode) > 0 Then txtCode_KeyDown vbKeyReturn, vbKeyShift

End Sub


Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtmcode
    Set PO_DESC = txtmdesc
    
    GoTop PR_itemmsetup
    MyLookup.Caption = "Item Units"
    MyLookup.FillGrid PR_itemmsetup, "mcode", "description", txtmcode.MaxLength
    MyLookup.Show 1
    
    If Len(txtmcode) > 0 Then txtmCode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_itemsetup, Me, txtCode, txtDesc, Para_Rs, "IC_GrnCnt", 3, "Itemcode", "description", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("CITY000001")
  SetToolBar(2) = chkRights("CITY000002")
  SetToolBar(3) = chkRights("CITY000003")
  SetToolBar(4) = chkRights("CITY000004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  

  PR_itemsetup.Open "Select * from ic_item where compcode = '" & Gs_compcode & "' Order By itemCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_itemmsetup.Open "Select * from IC_ItemUM  Order By MCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  PB_BlnkLoca = IIf(PR_itemsetup.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_itemsetup.Close
    PR_itemmsetup.Close
End Sub


Private Sub txtopeningpriceperitem_Validate(Cancel As Boolean)
If txtopeningpriceperitem <> "" Then
    txtopenvalue = Val(txtopeningpriceperitem) * (Val(txtopenqty) * Val(txtonofqty))
    txtwholeprice = Val(txtopeningpriceperitem) * Val(txtonofqty)
    txtnoofitems = Val(txtopenqty) * Val(txtonofqty)
End If
End Sub

Private Sub txtopeningunit_Click()
If txtopeningunit = "Single" Then
    txtonofqty = 1
    If txtopenvalue.Enabled Then txtopenvalue.SetFocus
Else
    If txtonofqty.Enabled Then txtonofqty.SetFocus
End If
End Sub

Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtmcode.SetFocus
End Sub
Private Sub txtitemtype_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If txtitemtype = "Single" Then
        txtnoofitem = 1
        If txtpurchaseprice.Enabled Then txtpurchaseprice.SetFocus
    Else
        txtnoofitem.SetFocus
    End If
    
  End If
End Sub
Private Sub txtnoofitem_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtpurchaseprice.SetFocus
End Sub
Private Function maxtranscode() As String

pr_dumy.Open "select max(itemcode) as transcode from ic_item where compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
pr_dumy.Close
End Function


Private Sub txtopenvalue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtopeningpriceperitem.SetFocus
End Sub

Private Sub txtopenvalue_LostFocus()
On Error Resume Next
If Val(txtopenvalue) > 0 And Val(txtopenqty) > 0 Then
    txtopeningpriceperitem = Val(txtopenvalue) / (Val(txtopenqty) * Val(txtonofqty))
    txtwholeprice = Val(txtopeningpriceperitem) * Val(txtonofqty)
    txtnoofitems = Val(txtopenqty) * Val(txtonofqty)
End If
End Sub

Private Sub txtpurchaseprice_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtsaleprice.SetFocus
End Sub

Private Sub txtpurchaseprice_LostFocus()
txttotalpurchaseprice = Val(txtpurchaseprice) * Val(txtnoofitem)
End Sub
Private Sub txtsaleprice_LostFocus()
txttotalsaleprice = Val(txtsaleprice) * Val(txtnoofitem)
End Sub

Private Sub txtsaleprice_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtweight.SetFocus
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtCode.Text = DoPad(txtCode.Text, txtCode.MaxLength)
      lb_found = MySeek(txtCode.Text, "Itemcode", PR_itemsetup)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtCode.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtCode.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                     ' txtLocation.Enabled = False
                      txtDesc.SetFocus
                   End If
                End If
            End Select
ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
End If
  End Sub
Private Sub txtmCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If KeyCode = vbKeyReturn Then
         
      txtmcode.Text = DoPad(txtmcode.Text, txtmcode.MaxLength)
      lb_found = MySeek(txtmcode.Text, "mcode", PR_itemmsetup)
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtmcode.SetFocus
                Else
                  txtnoofitem = PR_itemmsetup("U_Factor")
                  txtonofqty = PR_itemmsetup("U_Factor")
                  txtmdesc = PR_itemmsetup("Description")
                  txtomdesc = PR_itemmsetup("Description")
                End If
            
ElseIf KeyCode = vbKeyF12 Then
        Call Command1_Click
End If
  End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
    Else
      cmdLookup.Enabled = True
    End If
    
  If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_itemsetup, Me, txtCode, txtDesc, Para_Rs, "IC_GrnCnt", 3, "Itemcode", "description", 1, False, Toolbar1)
    End If
    
    If Mode = "A" Then
       txtCode = maxtranscode
       txtDesc.SetFocus
    End If
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
PB_BlnkLoca = False
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT into IC_Item(Compcode,Itemcode,description,mcode,mdesc,NoofItems,Purchaseprice,TotalPurchaseprice,Saleprice,TotalSalePrice,weight,openingqty,openingvalue,OpeningNoofitem,OpeningUnits,Priceperitem,Pricepeunits) VALUES ('" & Gs_compcode & "', '" & txtCode.Text & "','" & txtDesc.Text & "','" & txtmcode.Text & "','" & txtmdesc.Text & "'," & txtnoofitem.Text & "," & txtpurchaseprice.Text & "," & txttotalpurchaseprice.Text & "," & txtsaleprice.Text & "," & txttotalsaleprice.Text & "," & Val(txtweight) & "," & Val(txtopenqty) & "," & Val(txtopenvalue) & "," & Val(txtonofqty) & ",'" & Trim(txtmcode) & "'," & Val(txtopeningpriceperitem) & "," & Val(txtwholeprice) & ")"
           Case "E"
              gc_dbcon.Execute "UPDATE IC_Item SET description= '" & txtDesc.Text & "',mcode= '" & Trim(txtmcode.Text) & "',mdesc= '" & Trim(txtmdesc.Text) & "',noofitems= " & txtnoofitem.Text & ",purchaseprice= " & txtpurchaseprice.Text & ",totalpurchaseprice= " & txttotalpurchaseprice.Text & ",saleprice= " & txtsaleprice.Text & ",totalsaleprice= " & txttotalsaleprice.Text & ",weight = " & Val(txtweight) & " ,openingqty =" & Val(txtopenqty) & " ,openingvalue =" & Val(txtopenvalue) & ",OpeningNoofitem = " & Val(txtonofqty) & ",OpeningUnits = '" & Trim(txtmcode) & "',Priceperitem = " & Val(txtopeningpriceperitem) & ",Pricepeunits = " & Val(txtwholeprice) & "  WHERE  compcode = '" & Gs_compcode & "' and  Itemcode= '" & txtCode.Text & "'"
           Case "D"
           
           'check entry exist in any table
              
              pr_dumy.Open "Select * from Ic_trans where compcode = '" & Gs_compcode & "' and itemcode= '" & txtCode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
              If Not pr_dumy.EOF Then
                Call MsgBox("Record will not delete exist in inventory transaction", vbCritical)
                pr_dumy.Close
                gc_dbcon.CommitTrans
                
                Exit Sub
              End If
              pr_dumy.Close
              
           
              gc_dbcon.Execute "DELETE FROM IC_Item WHERE compcode = '" & Gs_compcode & "' and  Itemcode = '" & txtCode.Text & "'"
     End Select
     
gc_dbcon.CommitTrans
 
PR_itemsetup.Requery


'opening balance

If Mode = "A" Or Mode = "E" Then
    
        Dim ls_transcode As String
        Dim ls_Referencecode As String
        ls_Referencecode = "O" + txtCode
        ls_transcode = maxtranscodepayments
        gc_dbcon.Execute "Delete from Ic_Trans where compcode = '" & Gs_compcode & "' and referencecode = '" & ls_Referencecode & "' "
        
        If Val(txtopenqty) > 0 Then
    
        ls_sql = "INSERT into IC_Trans(Compcode, TransCode, SaleDate, PartyCode"
        ls_sql = ls_sql & " , ItemCode, Quantity, ItemRate, Amount, STaxRate, SaleTaxAmount"
        ls_sql = ls_sql & " , Remarks,itemtype,noofitems,Transtype,Referencecode)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','0000000000'"
        ls_sql = ls_sql & " ,'" & Format(Dtpopeningdate, "YYYY/MM/DD") & "','000000'"
        ls_sql = ls_sql & " ,'" & Trim(txtCode) & "'," & Val(txtopenqty) & ""
        ls_sql = ls_sql & " ," & Val(txtwholeprice) & "," & Val(txtopenvalue) & ""
        ls_sql = ls_sql & " ,0,0"
        ls_sql = ls_sql & " ,'Opening Balance'"
        ls_sql = ls_sql & " ,'" & Trim(txtopeningunit) & "'," & Val(txtnoofitems) & ",'P','" & ls_Referencecode & "')"
        gc_dbcon.Execute ls_sql
     
        End If
    'supplier
    pr_dumy.Open "select * from ic_supplier where compcode = '" & Gs_compcode & "' and Suppliercode = '000000'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            gc_dbcon.Execute "INSERT into IC_Supplier(Compcode,SupplierCode,Description,Address,NicNo,Phoneoffice,Phoneres,Mobile,AgrDate,CodeID,CountryCode,CityCode,TehseelCode,GlAccountNo,stregno,sectorcode,openingbalance) VALUES ('" & Gs_compcode & "','000000','Opening Balance','Nil','Nil','Nil','Nil','Nil','" & Format(Date, "YYYY/MM/DD") & "','S','','','','','','',0)"
        End If
    pr_dumy.Close
    
 End If
 If Mode = "A" Then
       txtCode = maxtranscode
       txtDesc.SetFocus
 End If
   

Exit Sub
LocalErr:
MsgBox Err.Description
End Sub

Private Sub SetVal()
    On Error Resume Next
     txtDesc = Trim(PR_itemsetup("description") & "")
     txtmcode = Trim(PR_itemsetup("mcode") & "")
     txtmdesc = Trim(PR_itemsetup("mdesc") & "")
     txtomdesc = Trim(PR_itemsetup("mdesc") & "")
     txtnoofitem = Val(0 & PR_itemsetup("Noofitems"))
     txtpurchaseprice = Val(0 & PR_itemsetup("Purchaseprice"))
     txttotalpurchaseprice = Val(0 & PR_itemsetup("TotalPurchaseprice"))
     txtsaleprice = Val(0 & PR_itemsetup("Saleprice"))
     txttotalsaleprice = Val(0 & PR_itemsetup("TotalSaleprice"))
     txtweight = Trim(PR_itemsetup("weight") & "")
     txtopenqty = Val(0 & PR_itemsetup("Openingqty"))
     txtopenvalue = Val(0 & PR_itemsetup("Openingvalue"))
     If Val(txtopenvalue) > 0 Then
     Call txtopenvalue_LostFocus
     End If
     txtonofqty = Val(0 & PR_itemsetup("OpeningNoofitem"))
     txtopeningunit = PR_itemsetup("OpeningUnits")
     
     txtopeningpriceperitem = Val(0 & PR_itemsetup("Priceperitem"))
     txtwholeprice = Val(0 & PR_itemsetup("Pricepeunits"))
     
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtCode.Text) = txtCode.MaxLength And txtDesc <> "" And txtmcode <> "" And txtmdesc <> "" And txtnoofitem <> "" And txtpurchaseprice <> "" And txttotalpurchaseprice <> "" And txtsaleprice <> "" And txttotalsaleprice <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  PR_itemsetup.Requery
End Sub

Private Sub txtweight_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtopenqty.SetFocus
End Sub
Private Sub txtopenqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtopenvalue.SetFocus
End Sub

