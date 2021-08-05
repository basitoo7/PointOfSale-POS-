VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form POS_FROM 
   Caption         =   "POINT OF SALE"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -1830
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTBARCODE 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   750
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   7
      Top             =   2100
      Width           =   2310
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexPOS 
      Height          =   6450
      Left            =   135
      TabIndex        =   6
      Top             =   1020
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   11377
      _Version        =   393216
      RowHeightMin    =   400
      GridColor       =   -2147483632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "INVOICE DATE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label TXTINVDATE 
      Alignment       =   1  'Right Justify
      Caption         =   "02-JAN-2008"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "INVOICE NO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label TXTINVNUMBER 
      Alignment       =   1  'Right Justify
      Caption         =   "0000000000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "GRAND TOTAL"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   1
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Label TXTINVTOTAL 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   0
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "FILE"
      Begin VB.Menu MNUEXIT 
         Caption         =   "EXIT"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu MNUDATA 
      Caption         =   "DATA"
      Begin VB.Menu MNUNEWINVOICE 
         Caption         =   "NEW BILL"
         Shortcut        =   ^N
      End
      Begin VB.Menu MNUREPORT 
         Caption         =   "REPORT"
      End
      Begin VB.Menu MNUSALESREP 
         Caption         =   "SALES REPORT IN CHART"
      End
   End
End
Attribute VB_Name = "POS_FROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public j, i
Public x, X2, lin
Public STARTDATe, ENDDATe As Date
Public RS As New ADODB.Recordset
Public myRS As New ADODB.Recordset
Public Cn As New ADODB.Connection
Public Answer
Public nF, Inp1, Z, b, P, c
Public Totrec
Public CY As Integer
Public CZ As Integer
Public Cz4, CX As Integer






Private Sub MSFlexPOS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then

   j = MSFlexPOS.MouseRow
   MsgBox "You right clicked row #" & j

End If

End Sub





Private Sub Form_Click()
LOADHEADER
Fill_Grid
Me.TXTBARCODE.Visible = False
End Sub

Private Sub mnuNewinvoice_Click()
Me.MSFlexPOS.Clear

LOADHEADER
Me.MSFlexPOS.Row = 1
Me.TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
Me.TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
Me.TXTBARCODE.Visible = True
Me.TXTBARCODE.TabIndex = 0

RS.Open "select max(invno) + 1 from pos_tbl", Cn, adOpenDynamic, adLockBatchOptimistic
Me.TXTINVNUMBER.Caption = Format(RS.Fields(0), "0")
RS.Close
Me.TXTBARCODE.Locked = False
End Sub

Private Sub MNUNEWINV_Click()

End Sub

Private Sub MNUREPORT_Click()
DataEnvironment1.Connection1.Open
DataEnvironment1.Command1 Me.TXTINVNUMBER.Caption
DataReport1.Show
End Sub

Private Sub MNUSALESREP_Click()
SALESREPORT_FORM.Show
End Sub

Private Sub MSFLEXPOS_EnterCell()
With MSFlexPOS
If .Col <> 1 Then
.CellBackColor = vbRed
Else
.CellBackColor = vbWhite
End If
End With
End Sub

Private Sub MSFLEXPOS_LeaveCell()
MSFlexPOS.CellBackColor = vbWhite
End Sub

Public Sub SetTxtBARCODE()   'put textbox over cell
        With TXTRSP
        .Top = MSFlexPOS.Top + MSFlexPOS.CellTop
        .Left = MSFlexPOS.Left + MSFlexPOS.CellLeft
        .Width = MSFlexPOS.CellWidth - 20
'        .Height = MSFlexPOS.CellHeight
       ' .Text = MSFlexPOS
        .Visible = True
        .SelStart = Len(.Text)
        .SetFocus
     End With
End Sub
    Public Sub BGLoss(anyControl As Control)
        If anyControl.Locked = False Then
            'anyControl.Backcolor = &H80000005
         anyControl.BackColor = vbYellow
        End If
    End Sub
    Public Sub HFocus(ByRef sText As Variant)
        With sText
            .SelStart = 0
            .SelLength = Len(sText.Text)
        End With
        
    End Sub



Public Sub DISAPEARME()
    Dim s
    For Each s In Controls
    If TypeOf s Is TextBox Then
    s.Visible = False
    End If
    Next
End Sub

Public Sub BGFocus(anyControl As Control)
        If anyControl.Locked = False Then
            'anyControl.Backcolor = &H80000013
            anyControl.BackColor = vbGreen
        End If
End Sub

Private Sub Form_Load()

Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Haleem\Desktop\temp\POS2122067282008\POS PROJECT\DATABASE\POS_DB.mdb;Persist Security Info=False"
Cn.CursorLocation = adUseClient


LOADHEADER

With MSFlexPOS
If .Col = 1 Then
'Me.TXTBARCODE.Text = UCase(Me.TXTBARCODE.Text)
Me.TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
Me.TXTBARCODE.Text = .TextMatrix(.Row, 1)
Me.TXTBARCODE.Visible = True
Me.TXTBARCODE.TabIndex = 0
End If
End With

'Me.TXTBARCODE.AddItem "0000000000001"
'Me.TXTBARCODE.AddItem "0000000000002"
'Me.TXTBARCODE.AddItem "0000000000003"
'Me.TXTBARCODE.AddItem "0000000000004"
'Me.TXTBARCODE.AddItem "0000000000005"
End Sub
Private Sub LOADHEADER()

With MSFlexPOS
'.RowHeight(0) = 500
    .Cols = 12
    .Rows = 2
    .ColWidth(0) = 800:
    .ColWidth(1) = 2500:
    .ColWidth(2) = 3300:
    .ColWidth(3) = 1500:
    .ColWidth(4) = 1100:
    .ColWidth(0) = 600:
    .ColWidth(7) = 1500:
    .ColWidth(10) = 1600:
    
    
'    .ColAlignmentFixed(0) = flexAlignLeftCenter
    '.ColAlignmentFixed(1) = flexAlignLeftCenter
    '.ColAlignmentFixed(2) = flexAlignRightCenter
    '.ColAlignmentFixed(3) = flexAlignRightCenter
    '.ColAlignmentFixed(4) = flexAlignRightCenter
    

    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignLeftCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColAlignment(6) = flexAlignRightCenter
    .ColAlignment(7) = flexAlignRightCenter
    .ColAlignment(8) = flexAlignRightCenter
    .ColAlignment(9) = flexAlignRightCenter
    .ColAlignment(10) = flexAlignRightCenter
    .ColAlignment(11) = flexAlignRightCenter
    

            .TextMatrix(0, 0) = "No"
            .TextMatrix(0, 1) = "Barcode"
            .TextMatrix(0, 2) = "Description"
            .TextMatrix(0, 3) = "Category"
            .TextMatrix(0, 4) = "U.O.M"
            .TextMatrix(0, 5) = "Sale Price"
            .TextMatrix(0, 6) = "Qty"
            .TextMatrix(0, 7) = "Sub Total"
            .TextMatrix(0, 8) = "Vat %"
            .TextMatrix(0, 9) = "Vat amount"
            .TextMatrix(0, 10) = "Line Total"
            .TextMatrix(0, 11) = "Row No"
        End With
        
        
    End Sub

Public Sub Fill_Data()
On Error Resume Next
RS.Open "SELECT COLOURNO FROM COLOUR_STATUS WHERE OBJECTNAME='BILLING GRID BACK COLOUR'", Cn, adOpenDynamic, adLockBatchOptimistic
MSFlexPOS.BackColor = RS.Fields(0)
RS.Close




LOADHEADER
RS.Open "SELECT * FROM BILLING_TBL WHERE INVOICENO = " & Me.TXTINVNUMBER.Caption & "", Cn, adOpenDynamic, adLockBatchOptimistic   'WHERE MAILTO = '" & MDI_FORM.StatusBar1.Panels(1).Text & "' order by ID", CN, adOpenDynamic, adLockBatchOptimistic

If RS.RecordCount = 0 Then
MsgBox "Records not found.."
Me.MSFlexPOS.Rows = 2
Me.MSFlexPOS.Cols = 2
Me.MSFlexPOS.Clear
RS.Close
Exit Sub
End If

Me.TXTINVNUMBER.Caption = RS.Fields![invoiceno]
Me.TXTINVDATE.Caption = RS.Fields![INVOICEDATE]


RS.MoveFirst
Me.MSFlexPOS.Rows = 2
Me.MSFlexPOS.Rows = RS.RecordCount + 1
Me.ProgressBar1.Max = Me.MSFlexPOS.Rows
'Me.Caption = "SEARCHING RECOR(S) " & CUR
CUR = 1
Do Until RS.EOF
'Me.Caption = "SEARCHING RECOR(S) " & CUR
           Me.MSFlexPOS.Row = CUR
           Me.MSFlexPOS.TextMatrix(CUR, 0) = CUR
           Me.MSFlexPOS.TextMatrix(CUR, 1) = RS.Fields![Category]
           Me.MSFlexPOS.TextMatrix(CUR, 2) = RS.Fields![PARTICULARS]
           Me.MSFlexPOS.TextMatrix(CUR, 3) = RS.Fields![ITEMSize]
           Me.MSFlexPOS.TextMatrix(CUR, 4) = RS.Fields![UOM]
           Me.MSFlexPOS.TextMatrix(CUR, 5) = Format(RS.Fields![RSP], "######0.00")
           Me.MSFlexPOS.TextMatrix(CUR, 6) = Format(RS.Fields![QTY], "######0.00")
           Me.MSFlexPOS.TextMatrix(CUR, 7) = Format(RS.Fields![Subtotal], "######0.00")
           Me.MSFlexPOS.TextMatrix(CUR, 8) = Format(RS.Fields![VATPERCENT], "######0.00")
           Me.MSFlexPOS.TextMatrix(CUR, 9) = Format(RS.Fields![VATAMOUNT], "######0.00")
           Me.MSFlexPOS.TextMatrix(CUR, 10) = Format(RS.Fields![LINETOTAL], "######0.00")
           Me.MSFlexPOS.TextMatrix(CUR, 11) = RS.Fields![ID]
        'CUR = CUR + 1
RS.MoveNext
CUR = CUR + 1
'ShowProgressInStatusBar True
Me.ProgressBar1.Value = CUR
Loop
RS.Close
'MAKESUM
Me.ProgressBar1.Value = 0


'RS.Open "SELECT COLOURNO FROM COLOUR_STATUS WHERE OBJECTNAME='BILLING GRID LINE COLOUR'", CN, adOpenDynamic, adLockBatchOptimistic
RS.Open "SELECT COLOURNO FROM COLOUR_STATUS WHERE OBJECTNAME='BILLING FORM LINE COLOUR'", Cn, adOpenDynamic, adLockBatchOptimistic


Me.MSFlexPOS.Col = 1
Me.MSFlexPOS.Row = 1

Me.MSFlexPOS.ForeColorFixed = vbRed
    Do Until MSFlexPOS.Row = MSFlexPOS.Rows - 1
            
            If Me.MSFlexPOS.Row = 0 Then
                For K = 0 To Me.MSFlexPOS.Cols
                Me.MSFlexPOS.CellFontBold = True
                Me.MSFlexPOS.CellBackColor = vbBlue
                Me.MSFlexPOS.Col = K
                Next K
                
            End If

            
        MSFlexPOS.Row = MSFlexPOS.Row + 1
        For c = 1 To MSFlexPOS.Cols - 1
            MSFlexPOS.Col = c
             
            'MSFlexPOS.CellBackColor = vbYellow
            Me.MSFlexPOS.CellBackColor = RS.Fields(0)
        Next
        
        If Me.MSFlexPOS.Row <> Me.MSFlexPOS.Rows - 1 Then
        MSFlexPOS.Row = MSFlexPOS.Row + 1
        End If
        
    Loop
    MSFlexPOS.Col = 1
    
RS.Close
End Sub

Private Sub MNUEXIT_Click()
Unload Me
End Sub

Private Sub MNUSEARCHDATA_Click()
Fill_Data
End Sub

Private Sub MSFlexPOS_Click()
With MSFlexPOS
If .Col = 1 Then
'Me.TXTBARCODE.Text = UCase(Me.TXTBARCODE.Text)
Me.TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
Me.TXTBARCODE.Text = .TextMatrix(.Row, 1)
Me.TXTBARCODE.Visible = True
Me.TXTBARCODE.SetFocus
End If
End With
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TXTBARCODE_DblClick()
With MSFlexPOS
If .Col = 1 Then
          Me.TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
             Me.TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
               Me.TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
                'Me.TXTPARTICULARS.Visible = False
                Me.TXTBARCODE.Visible = True
                Me.TXTBARCODE.SetFocus
                End If
    End With
End Sub



Public Sub HighLight()
        With Screen.ActiveForm
            If (TypeOf .ActiveControl Is TextBox) Then
                .ActiveControl.SelStart = 0
                .ActiveControl.SelLength = Len(.ActiveControl)
    '         .ActiveControl.Backcolor = &H80000013
            .ActiveControl.BackColor = vbYellow
            End If
        End With
End Sub

Private Sub TXTBARCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then



If Len(Me.TXTBARCODE.Text) >= 1 Then
Me.TXTBARCODE.Text = Format(Me.TXTBARCODE.Text, "0000000000000")
End If



RS.Open "SELECT * FROM item_master_TBL WHERE BARCODE = '" & Me.TXTBARCODE.Text & "'", Cn, adOpenDynamic, adLockBatchOptimistic

If RS.EOF = False Then

Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 0) = Me.MSFlexPOS.Row
               Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 1) = RS.Fields![barcode]
               Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 2) = RS.Fields![Description]
               Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 3) = RS.Fields![Category]
               Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 4) = RS.Fields![UOM]
               Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 5) = RS.Fields![COST]
               Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 6) = "1.00"
               Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 8) = "0"
With Me.MSFlexPOS
      .TextMatrix(.Row, 7) = .TextMatrix(.Row, 5) * .TextMatrix(.Row, 6)
       .TextMatrix(.Row, 10) = .TextMatrix(.Row, 7) + .TextMatrix(.Row, 8)
End With
      
If Me.MSFlexPOS.Row = Me.MSFlexPOS.Rows - 1 Then
MSFlexPOS.Col = 1
Me.MSFlexPOS.Row = Me.MSFlexPOS.Rows - 1

With Me.MSFlexPOS

If .TextMatrix(.Row, 11) = "" Then
Cn.Execute "INSERT INTO POS_TBL(INVNO,INVDATE,BARCODE,DESCRIPTION,CATEGORY,UOM,RSP,QTY,STOTAL,VATPER,VATAMT,LINETOTAL)VALUES('" & _
Me.TXTINVNUMBER.Caption & "','" & _
Me.TXTINVDATE.Caption & "','" & _
    .TextMatrix(.Row, 1) & "','" & _
    .TextMatrix(.Row, 2) & "','" & _
    .TextMatrix(.Row, 3) & "','" & _
    .TextMatrix(.Row, 4) & "'," & _
    .TextMatrix(.Row, 5) & "," & _
    .TextMatrix(.Row, 6) & "," & _
    .TextMatrix(.Row, 7) & "," & _
Val(.TextMatrix(.Row, 8)) & "," & _
Val(.TextMatrix(.Row, 9)) & "," & _
Val(.TextMatrix(.Row, 10)) & ")"
End If

End With

myRS.Open "SELECT SUM(LINETOTAL) FROM POS_TBL WHERE INVNO = " & Me.TXTINVNUMBER.Caption & "", Cn, adOpenDynamic, adLockBatchOptimistic

Me.TXTINVTOTAL.Caption = myRS.Fields(0)

myRS.Close


Me.Caption = ""
Me.MSFlexPOS.Rows = Me.MSFlexPOS.Rows + 1
Me.MSFlexPOS.Row = Me.MSFlexPOS.Rows - 1
TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
 TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
 TXTBARCODE.Visible = True
 ElseIf Me.MSFlexPOS.Row < Me.MSFlexPOS.Rows - 1 Then
 
Me.Caption = "edit"
Me.MSFlexPOS.Row = Me.MSFlexPOS.Rows - 1
TXTBARCODE.Move MSFlexPOS.Left + MSFlexPOS.CellLeft - CX, MSFlexPOS.Top + MSFlexPOS.CellTop - CY, MSFlexPOS.CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
 TXTBARCODE.Text = MSFlexPOS.TextMatrix(MSFlexPOS.Row, 1)
 End If
 RS.Close
Else
MsgBox "BARCODE IS NOT VALID", vbExclamation, "Error"
RS.Close
End If


'Me.TXTEVENTDATE.Text = UCase(Me.TXTPARTICULARS.Text)
With Me.MSFlexPOS
'.TextMatrix(.Row, 1) = Me.TXTEVENTDATE.Text
Me.TXTBARCODE.Visible = False
.Col = 1
'Me.TXTBARCODE.SetFocus

Me.TXTBARCODE.Visible = True
Me.TXTBARCODE.SetFocus
End With
End If
End Sub




Public Sub Fill_Grid()
Dim i
'On Error Resume Next
i = 1
'RS.Open "SELECT * FROM COLOUR_STATUS", CN, adOpenDynamic, adLockBatchOptimistic
RS.Open "select max(invno) from pos_tbl", Cn, adOpenDynamic, adLockBatchOptimistic
Me.TXTINVNUMBER.Caption = Format(RS.Fields(0), "0")
RS.Close

RS.Open "SELECT * FROM POS_TBL WHERE INVNO = " & Me.TXTINVNUMBER.Caption & "", Cn, adOpenDynamic, adLockBatchOptimistic

 Me.MSFlexPOS.Rows = RS.RecordCount

While Not RS.EOF

Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 0) = Me.MSFlexPOS.Row
Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 1) = RS.Fields(3)
Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 2) = RS.Fields(4)
Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 3) = RS.Fields(5)
Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 4) = RS.Fields(6)

Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 5) = RS.Fields(8)
Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 6) = RS.Fields(9)


Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 7) = RS.Fields(10)
Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 8) = RS.Fields(11)


Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 9) = RS.Fields(12)
Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 10) = RS.Fields(13)
Me.MSFlexPOS.TextMatrix(Me.MSFlexPOS.Row, 11) = RS.Fields(0)

If Me.MSFlexPOS.Row = Me.MSFlexPOS.Rows - 1 Then
Else
MSFlexPOS.Row = MSFlexPOS.Row + 1
End If


'MSFlexPOS.CellFontBold = True
'MSFlexPOS.CellBackColor = RS.Fields(1)
RS.MoveNext
 'MSFLEXPOS.Row = MSFLEXPOS.Row + 1

 i = i + 1
Wend
RS.Close

End Sub




