VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmicreport10 
   Caption         =   "List of Items"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports10.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkSticker 
      Caption         =   "New Sticker "
      Height          =   375
      Left            =   2640
      TabIndex        =   51
      Top             =   4560
      Width           =   1605
   End
   Begin VB.CheckBox ChkDate 
      Caption         =   "Date"
      Height          =   375
      Left            =   315
      TabIndex        =   50
      Top             =   2835
      Width           =   690
   End
   Begin VB.TextBox txtToDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   2985
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2115
      Width           =   3510
   End
   Begin VB.CommandButton Command5 
      Height          =   315
      Left            =   2655
      Picture         =   "frmicreports10.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2100
      Width           =   315
   End
   Begin VB.TextBox txtToCode 
      Height          =   315
      Left            =   2010
      TabIndex        =   46
      Top             =   2100
      Width           =   645
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   2010
      TabIndex        =   44
      Top             =   1725
      Width           =   645
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   2985
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1770
      Width           =   3510
   End
   Begin VB.CommandButton CmdLooup 
      Height          =   315
      Left            =   2655
      Picture         =   "frmicreports10.frx":047C
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1740
      Width           =   315
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   3015
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2490
      Width           =   3510
   End
   Begin VB.CommandButton CmdUserInfo 
      Height          =   315
      Left            =   2685
      Picture         =   "frmicreports10.frx":05EE
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2475
      Width           =   345
   End
   Begin VB.TextBox txtuserid 
      Height          =   315
      Left            =   2025
      TabIndex        =   38
      Top             =   2460
      Width           =   630
   End
   Begin VB.CheckBox ChkWasNow 
      Caption         =   "Was ,Now"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   4560
      Width           =   1605
   End
   Begin VB.CheckBox ChkSaleCost 
      Caption         =   "Only Sale Cost"
      Height          =   375
      Left            =   2640
      TabIndex        =   30
      Top             =   4200
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Description"
            TextSave        =   "Description"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   105833
            MinWidth        =   105833
         EndProperty
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
   End
   Begin VB.Frame Frame1 
      Height          =   5085
      Left            =   30
      TabIndex        =   3
      Top             =   -45
      Width           =   6675
      Begin VB.CheckBox chkratediff 
         Caption         =   "Sale Rate Difference With Net Rate Report"
         Height          =   390
         Left            =   240
         TabIndex        =   27
         Top             =   3825
         Width           =   4005
      End
      Begin VB.CheckBox chksupplierwise 
         Caption         =   "Display Items Supplier Wise"
         Height          =   345
         Left            =   195
         TabIndex        =   10
         Top             =   4200
         Width           =   2520
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   75
         Top             =   585
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   3840
         Left            =   0
         TabIndex        =   4
         Top             =   -60
         Width           =   6675
         Begin VB.CheckBox ChkTime 
            Caption         =   "Time"
            Height          =   375
            Left            =   285
            TabIndex        =   37
            Top             =   3375
            Width           =   690
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2595
            Picture         =   "frmicreports10.frx":0760
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   225
            Width           =   315
         End
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2955
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   225
            Width           =   3510
         End
         Begin VB.TextBox txtdeptcode 
            Height          =   315
            Left            =   2010
            MaxLength       =   3
            TabIndex        =   20
            Top             =   210
            Width           =   555
         End
         Begin VB.TextBox txtcatcode 
            Height          =   315
            Left            =   1995
            TabIndex        =   19
            Top             =   615
            Width           =   555
         End
         Begin VB.TextBox txtcatdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2955
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   600
            Width           =   3510
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2595
            Picture         =   "frmicreports10.frx":08D2
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   600
            Width           =   315
         End
         Begin VB.TextBox txtsubcatcode 
            Height          =   315
            Left            =   1980
            TabIndex        =   16
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox txtsubcatdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2970
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1020
            Width           =   3510
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2595
            Picture         =   "frmicreports10.frx":0A44
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1005
            Width           =   315
         End
         Begin VB.TextBox txtSuppliercode 
            Height          =   315
            Left            =   1965
            TabIndex        =   13
            Top             =   1395
            Width           =   645
         End
         Begin VB.TextBox txtSupplierdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   2955
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1440
            Width           =   3510
         End
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   2610
            Picture         =   "frmicreports10.frx":0BB6
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1380
            Width           =   315
         End
         Begin MSComCtl2.DTPicker DTPto 
            Height          =   315
            Left            =   4290
            TabIndex        =   7
            Top             =   2970
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100270083
            CurrentDate     =   41216
         End
         Begin MSComCtl2.DTPicker DTPfrom 
            Height          =   315
            Left            =   1950
            TabIndex        =   32
            Top             =   2955
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100270083
            CurrentDate     =   41216
         End
         Begin MSComCtl2.DTPicker DTPTimeFrom 
            Height          =   315
            Left            =   1965
            TabIndex        =   33
            Top             =   3450
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   100270082
            CurrentDate     =   41216
         End
         Begin MSComCtl2.DTPicker DTPTimeTo 
            Height          =   315
            Left            =   4335
            TabIndex        =   34
            Top             =   3450
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   100270082
            CurrentDate     =   41216
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "To Item Code :"
            Height          =   210
            Left            =   915
            TabIndex        =   49
            Top             =   2250
            Width           =   1020
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "From Item Code :"
            Height          =   210
            Left            =   705
            TabIndex        =   45
            Top             =   1845
            Width           =   1200
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "User Name :"
            Height          =   210
            Left            =   1020
            TabIndex        =   41
            Top             =   2595
            Width           =   885
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "To Time :"
            Height          =   210
            Left            =   3630
            TabIndex        =   36
            Top             =   3525
            Width           =   645
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "From Time :"
            Height          =   210
            Left            =   1035
            TabIndex        =   35
            Top             =   3495
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Department Code :"
            Height          =   210
            Left            =   585
            TabIndex        =   26
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Category Code :"
            Height          =   210
            Left            =   735
            TabIndex        =   25
            Top             =   645
            Width           =   1170
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sub Cat Code :"
            Height          =   210
            Left            =   780
            TabIndex        =   24
            Top             =   1065
            Width           =   1080
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Supplier Code :"
            Height          =   210
            Left            =   765
            TabIndex        =   23
            Top             =   1455
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   3555
            TabIndex        =   9
            Top             =   3015
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            Height          =   210
            Left            =   1035
            TabIndex        =   8
            Top             =   3015
            Width           =   825
         End
         Begin VB.Label txtselectiveitem 
            Height          =   300
            Left            =   285
            TabIndex        =   6
            Top             =   780
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Label txtselectiveaccount 
            Height          =   300
            Left            =   3120
            TabIndex        =   5
            Top             =   780
            Visible         =   0   'False
            Width           =   2520
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   5535
         TabIndex        =   2
         Top             =   4170
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4380
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Top             =   4185
         Width           =   1050
      End
   End
   Begin VB.CommandButton CmdChangPer 
      Caption         =   "Click for Cahange Rate"
      Height          =   375
      Left            =   1080
      TabIndex        =   28
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox textNewPer 
      Height          =   315
      Left            =   3120
      TabIndex        =   29
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "frmicreport10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Dim PR_itemsetup As New Recordset
Public codeid As String
Public Reporttype As String
Dim ls_sql As String
Dim SQLQueryDiscPer As String





Private Sub ChkDate_Click()
If ChkDate.Value = 1 Then
    DTPfrom.Enabled = True
    DTPto.Enabled = True
Else
    DTPfrom.Enabled = False
    DTPto.Enabled = False
End If
End Sub

Private Sub Chktime_Click()
If ChkTime.Value = 1 Then
   DTPTimeFrom.Enabled = True
   DTPTimeTo.Enabled = True
Else
   DTPTimeFrom.Enabled = False
   DTPTimeTo.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdChangPer_Click()
SQLQueryDiscPer = ""
If textNewPer.Text = "" Then
   MsgBox ("Enter Persontage ....:"), vbCritical
Else
  
        If txtdeptcode <> "" Then
           SQLQueryDiscPer = "catcode='" & Trim(txtdeptcode.Text) & "'"
         End If
        If txtcatcode <> "" Then
           SQLQueryDiscPer = SQLQueryDiscPer & "and ClassId='" & Trim(txtcatcode.Text) & "'"
         End If
        If txtsubcatcode <> "" Then
           SQLQueryDiscPer = SQLQueryDiscPer & "and PackCode='" & Trim(txtsubcatcode.Text) & "'"
         End If
        If txtSuppliercode <> "" Then
           SQLQueryDiscPer = SQLQueryDiscPer & "and manucode='" & Trim(txtSuppliercode.Text) & "'"
        End If
      
        If SQLQueryDiscPer = "" Then
           
           MsgBox ("Please Enter Any Oprion ...:"), vbCritical
         Else
          
         ' gc_dbcon.Execute ("update ic_item set fxdper=Round(((SaleCost) - (PurchaseCost)) / ((PurchaseCost) / 100), 2) Where " & SQLQueryDiscPer & "")
          
          
         ' gc_dbcon.Execute ("update ic_item set Salecost = (Purchasecost + Round(((purchasecost / 100) * " & Val(textNewPer) & "), 0)) Where " & SQLQueryDiscPer & "")
          
        '  gc_dbcon.Execute ("update ic_item set Salecost =  Round(Salecost,0) Where " & SQLQueryDiscPer & "")
          
          
         
         gc_dbcon.Execute ("update ic_item set RoundCost=  right(SaleCost,6) Where " & SQLQueryDiscPer & "")
          
         gc_dbcon.Execute ("update ic_item set SaleCost = SaleCost -1 Where RoundCost = 1 and DiscAmt <=0  and " & SQLQueryDiscPer & "")
         
         gc_dbcon.Execute ("update ic_item set SaleCost = SaleCost -2 Where RoundCost = 2 and DiscAmt <=0 and " & SQLQueryDiscPer & "")
         
          gc_dbcon.Execute ("update ic_item set SaleCost = SaleCost + 2 Where RoundCost = 3 and DiscAmt <=0 and " & SQLQueryDiscPer & "")
          
            gc_dbcon.Execute ("update ic_item set SaleCost = SaleCost + 1 Where RoundCost = 4 and DiscAmt <=0 and " & SQLQueryDiscPer & "")
          
          
          
           gc_dbcon.Execute ("update ic_item set SaleCost = SaleCost -1 Where RoundCost = 6 and DiscAmt <=0 and " & SQLQueryDiscPer & "")
         
         gc_dbcon.Execute ("update ic_item set SaleCost = SaleCost -2 Where RoundCost = 7 and DiscAmt <=0 and " & SQLQueryDiscPer & "")
         
          gc_dbcon.Execute ("update ic_item set SaleCost = SaleCost + 2 Where RoundCost = 8 and DiscAmt <=0 and " & SQLQueryDiscPer & "")
          
            gc_dbcon.Execute ("update ic_item set SaleCost = SaleCost + 1 Where RoundCost = 9 and DiscAmt <=0  and " & SQLQueryDiscPer & "")
          
          
          
          
          'gc_dbcon.Execute ("update ic_item set fxdper=Round(((SaleCost) - (PurchaseCost)) / ((PurchaseCost) / 100), 2) where manucode='" & Trim(txtSuppliercode.Text) & "' and catcode<> 29")
          'gc_dbcon.Execute ("update ic_item set Salecost = (Purchasecost + Round(((purchasecost / 100) * " & Val(textNewPer) & "), 0)) where fxdPer >= 10 and manucode='" & Trim(txtSuppliercode.Text) & "' and catcode<>29")
          'gc_dbcon.Execute ("update ic_item set Salecost =  Round(Salecost,0) where manucode='" & Trim(txtSuppliercode.Text) & "' and catcode<> 29")
       
            
       
       
          'gc_dbcon.Execute ("update ic_item set fxdper=Round(((SaleCost) - (PurchaseCost)) / ((PurchaseCost) / 100), 2) where manucode='" & Trim(txtSuppliercode.Text) & "' and catcode='" & Trim(txtdeptcode.Text) & "' and ClassId='" & Trim(txtcatcode.Text) & "' and PackCode='" & Trim(txtsubcatcode.Text) & "'")
          'gc_dbcon.Execute ("update ic_item set Salecost = (Purchasecost + Round(((purchasecost / 100) * " & Val(textNewPer) & "), 0)) where fxdPer >= 10 and manucode='" & Trim(txtSuppliercode.Text) & "' and catcode='" & Trim(txtdeptcode.Text) & "' and ClassId='" & Trim(txtcatcode.Text) & "' and PackCode='" & Trim(txtsubcatcode.Text) & "'")
          'gc_dbcon.Execute ("update ic_item set Salecost =  Round(Salecost,0) where manucode='" & Trim(txtSuppliercode.Text) & "' and catcode='" & Trim(txtdeptcode.Text) & "' and ClassId='" & Trim(txtcatcode.Text) & "' and PackCode='" & Trim(txtsubcatcode.Text) & "'")
        End If
      
      
   MsgBox ("Process Successfull .....:"), vbCritical
End If


End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
    
If Me.Caption = "List of Items" Then
    
    With crrpt
    
     If chksupplierwise.Value = 1 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_Itemssupplierwise.RPT"
     
     ElseIf chksupplierwise.Value = 1 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemsRateDiff.RPT"
     ElseIf ChkSaleCost.Value = 1 Then
      '.ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemsSaleCost.rpt"
      .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemsSupplierWiseSaleCost.rpt"
     Else
      .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_Items.RPT"
     End If
        
        .WindowTitle = "Company Items"
        .SelectionFormula = "{Ic_item.Compcode} = '" & Gs_compcode & "'"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Company Items'"
        

        
        
        .SelectionFormula = ""
        
        .SelectionFormula = "{Ic_Item.compcode} = '" & Gs_compcode & "'"
        
        If txtdeptcode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {Ic_Item.Catcode} = '" & Trim(txtdeptcode) & "'"
        End If
        
         If ChkDate.Value = 1 Then
              .SelectionFormula = .SelectionFormula & " and {Ic_Item.AddDateTime} >= Date(" & DTPfrom.Year & "," & DTPfrom.Month & "," & DTPfrom.Day & ") AND {Ic_Item.AddDateTime} <= Date(" & DTPto.Year & "," & DTPto.Month & "," & DTPto.Day & ")"
        End If
        If txtcatcode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {Ic_Item.classid} = '" & Trim(txtcatcode) & "'"
        End If
        
        If txtCode <> "" And txtToCode = "" Then
        .SelectionFormula = .SelectionFormula & "  and {Ic_Item.itemCode} = '" & Trim(txtCode) & "'"
        End If
         
       If txtCode <> "" And txtToCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and Val({Ic_Item.itemCode}) >= " & Val(txtCode) & " and Val({Ic_Item.itemCode}) <= " & Val(txtToCode) & ""
        End If
         
         
        If txtsubcatcode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {Ic_Item.packcode} = '" & Trim(txtsubcatcode) & "'"
        End If
        
        If txtSuppliercode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {Ic_Item.manucode} = '" & Trim(txtSuppliercode) & "'"
        End If
        If chkratediff.Value = 1 Then
        .SelectionFormula = .SelectionFormula & "  and {Ic_Item.AvgRate}> {Ic_Item.SaleCost}"
        End If
        
               
        .Connect = "DNS=Censoft;UID=Sa"
        
        
        
        .Action = 1
       
       
    End With
Else
 With crrpt
         If Me.Caption = "Purchase List Rate Change Periodic" Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemslogP.rpt"
             .Formulas(1) = "ReportName = 'New Purchase Rate Item List'"
        Else
            .Formulas(1) = "ReportName = 'New Sale Rate Item List'"
         
         
          If ChkSticker.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemslogWasNow.rpt"
             '  .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemsWasNow.rpt"
            '   .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemsWasNow2.rpt"
          ElseIf ChkWasNow.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemslogWasNow.rpt"
             '  .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemsWasNow.rpt"
            '   .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemsWasNow2.rpt"
          Else
            '.ReportFileName = App.Path & Gs_ICRepoPath & "\IC_Itemslog.rpt"
             .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemslogD.rpt"
          End If
        End If
        
        .SQLQuery = ""
        .SQLQuery = "SELECT Ic_ItemSetupLog.Transdate,Ic_ItemSetupLog.PPCost,Ic_ItemSetupLog.CPCost, Ic_ItemSetupLog.PCost, Ic_ItemSetupLog.CCost, IC_Item.ItemCode, IC_Item.CustomCode, IC_Item.Description,AddUser, Ic_ItemSetupLog.CDiscAmt"
        .SQLQuery = .SQLQuery & " IC_Item.CatCode, IC_ItemCategory.Description AS CateDesc FROM Ic_ItemSetupLog Ic_ItemSetupLog LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Item IC_Item ON Ic_ItemSetupLog.Compcode = IC_Item.Compcode AND Ic_ItemSetupLog.ItemCode = IC_Item.ItemCode INNER JOIN"
        .SQLQuery = .SQLQuery & " IC_ItemCategory IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode"
        .SQLQuery = .SQLQuery & " where IC_Item.compcode = '" & Gs_compcode & "' "
        .SQLQuery = .SQLQuery & " and Ic_ItemSetupLog.Transdate >= '" & Format(DTPfrom.Value, "YYYY/MM/DD") & "' and Ic_ItemSetupLog.Transdate <= '" & Format(DTPto.Value, "YYYY/MM/DD") & "' and Ic_ItemSetupLog.CDiscAmt > 0 "
        
         If ChkTime.Value = 1 Then
            .SQLQuery = .SQLQuery & " and Ic_ItemSetupLog.Transdate >= '" & Format(DTPTimeFrom.Value, "HH:MM:SS") & "' and Ic_ItemSetupLog.Transdate <= '" & Format(DTPTimeTo.Value, "HH:MM:SS") & "' and Ic_ItemSetupLog.CDiscAmt > 0 "
         End If
               
         
         If Me.Caption = "Purchase List Rate Change Periodic" Then
           .SQLQuery = .SQLQuery & " and Ic_ItemSetupLog.PPCost+Ic_ItemSetupLog.CPCost <> 0"
           Else
            .SQLQuery = .SQLQuery & " and Ic_ItemSetupLog.PCost+Ic_ItemSetupLog.CCost <> 0"
          
         End If
        
        If txtuserid <> "" Then
          .SQLQuery = .SQLQuery & " and Ic_ItemSetupLog.AddUser ='" & LCase(Trim(txtuserid)) & "'"
        End If
        
        If txtdeptcode <> "" Then
            .SQLQuery = .SQLQuery & " and IC_Item.catcode = '" & txtdeptcode & "'"
        End If
        
        
        If txtcatcode <> "" Then
         .SQLQuery = .SQLQuery & "   and Ic_Item.classid = '" & Trim(txtcatcode) & "'"
        End If
        
         
        If txtsubcatcode <> "" Then
         .SQLQuery = .SQLQuery & "   and Ic_Item.packcode = '" & Trim(txtsubcatcode) & "'"
        End If
        
        If txtSuppliercode <> "" Then
         .SQLQuery = .SQLQuery & "   and Ic_Item.manucode = '" & Trim(txtSuppliercode) & "'"
        End If
        
          If txtCode <> "" Then
            .SQLQuery = .SQLQuery & "  and Ic_Item.ItemCode = '" & Trim(txtCode) & "'"
           End If
        
        
    
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
       
        .Formulas(2) = "Period = '" & "From " & DTPfrom & " to " & DTPto & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       
    End With
End If
    
Exit Sub

LocalErr:

Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub CmdLooup_Click()
  Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCode
    Set PO_DESC = txtDescription
    Gs_SQL = "Select IC_Item.ItemCode,   IC_Item.Description, IC_ItemCategory.Description as Category,IC_Item.SaleCost,StockS,StockG,isnull(StockS,0)+isnull(StockG,0) as TotalStock from IC_Item left outer join IC_ItemCategory on IC_Item.compcode = IC_ItemCategory.compcode and   IC_Item.catcode = IC_ItemCategory.catcode "
    Gs_FindFld = "IC_Item.Description"
    Gs_OrderBy = "Order by IC_Item.Description"
    Gs_OtherPara = " where IC_Item.compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1

    If Len(txtCode) > 0 Then txtCode_KeyDown vbKeyReturn, vbKeyShift
End Sub



Private Sub Command5_Click()
  Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCode
    Set PO_DESC = txtDescription
    Gs_SQL = "Select IC_Item.ItemCode,   IC_Item.Description, IC_ItemCategory.Description as Category,IC_Item.SaleCost,StockS,StockG,isnull(StockS,0)+isnull(StockG,0) as TotalStock from IC_Item left outer join IC_ItemCategory on IC_Item.compcode = IC_ItemCategory.compcode and   IC_Item.catcode = IC_ItemCategory.catcode "
    Gs_FindFld = "IC_Item.Description"
    Gs_OrderBy = "Order by IC_Item.Description"
    Gs_OtherPara = " where IC_Item.compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1

    If Len(txtToCode) > 0 Then txtCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And txtCode <> "" Then
         
      txtCode.Text = DoPad(txtCode.Text, txtCode.MaxLength)
      If PR_itemsetup.State = 1 Then PR_itemsetup.Close
      PR_itemsetup.Open "Select * from ic_item where compcode = '" & Gs_compcode & "' and itemcode = '" & txtCode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
      If PR_itemsetup.RecordCount > 0 Then
         txtDescription = PR_itemsetup("Description")
      Else
        MsgBox ("Item Not found ......")
        txtCode.SetFocus
      End If
        PR_itemsetup.Close
ElseIf KeyCode = vbKeyReturn And Trim(txtCode) = "" Then
        Call CmdLooup_Click
End If
End Sub
Private Sub txttoCode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And txtToCode <> "" Then
         
      txtToCode.Text = DoPad(txtToCode.Text, txtToCode.MaxLength)
      If PR_itemsetup.State = 1 Then PR_itemsetup.Close
      PR_itemsetup.Open "Select * from ic_item where compcode = '" & Gs_compcode & "' and itemcode = '" & txtToCode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
      If PR_itemsetup.RecordCount > 0 Then
         txtToDescription = PR_itemsetup("Description")
      Else
        MsgBox ("Item Not found ......")
        txtCode.SetFocus
      End If
        PR_itemsetup.Close
ElseIf KeyCode = vbKeyReturn And Trim(txtCode) = "" Then
        Call CmdLooup_Click
End If
End Sub

Private Sub CmdUserInfo_Click()
    Set PO_CODE = Nothing
    Set PO_DESC = Nothing
    Set PO_AnyForm = Nothing
    
    Set PO_AnyForm = Me
    Set PO_CODE = txtuserid
    Set PO_DESC = Text4
    
    Gs_SQL = "Select UserID, UserName from SyUsers"
    Gs_FindFld = "UserName"
    Gs_OrderBy = "Order by UserName"
    'Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "User Information"
    MyLookupOLDB.Show 1
    
   cmdGenerate.SetFocus






End Sub



Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtdeptcode
    Set PO_DESC = txtdesc
    Gs_SQL = "Select CatCode,   Description from IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    
    If txtdeptcode <> "" Then Call txtdeptcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub


Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcatcode
    Set PO_DESC = txtcatdesc
    Gs_SQL = "Select ClassCode,   Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "'"
    MyLookupOLDB.Caption = "Categories"
    MyLookupOLDB.Show 1
    
    If txtcatcode <> "" Then Call txtcatcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsubcatcode
    Set PO_DESC = txtsubcatdesc
    Gs_SQL = "Select PackCode,   Description from IC_ItemPacking "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and subcode = '" & txtcatcode & "' and deptcode = '" & txtdeptcode & "' "
    MyLookupOLDB.Caption = "Sub Categories"
    MyLookupOLDB.Show 1
    
    If txtsubcatcode <> "" Then Call txtsubcatcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command4_Click()
      Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSuppliercode
    Set PO_DESC = txtSupplierdesc
    Gs_SQL = "Select SupplierCode,   Description from IC_Supplier "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Supplier"
    MyLookupOLDB.Show 1

    
    If txtSuppliercode <> "" Then Call txtSuppliercode_KeyDown(vbKeyReturn, vbKeyShift)


End Sub



Private Sub Form_Load()
If Trim(Gc_UserName) = "Administrator" Then
    CmdChangPer.Enabled = True
    textNewPer.Enabled = True
    ChkSaleCost.Enabled = True
End If
DTPfrom = Date
DTPto = Date
DTPTimeFrom = Time
DTPTimeTo = Time
If Me.Caption = "List of Items" Then
   DTPfrom.Enabled = False
   DTPto.Enabled = False
End If
End Sub

Private Sub txtcatcode_Change()
If txtcatcode = "" Then
txtcatdesc = ""
End If
End Sub

Private Sub txtcatcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtcatcode) <> "" And KeyCode = vbKeyReturn Then
        txtcatcode = DoPad(txtcatcode, 3)
       If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtcatcode & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Category Code not found !!!", vbCritical)
            txtcatcode = ""
            txtcatdesc = ""
            txtcatcode.SetFocus
        Else
            txtcatdesc = pr_dumy("Description")
             If txtsubcatcode.Enabled Then txtsubcatcode.SetFocus
           
        End If
        pr_dumy.Close
        
ElseIf Trim(txtcatcode) = "" And KeyCode = vbKeyReturn Then
        txtcatcode = ""
        txtcatdesc = ""
        Command2_Click
End If
End Sub

Private Sub txtdeptcode_Change()
If txtdeptcode = "" Then
txtdesc = ""
End If
End Sub

Private Sub txtdeptcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtdeptcode <> "" And KeyCode = vbKeyReturn Then
    txtdeptcode = DoPad(txtdeptcode, txtdeptcode.MaxLength)
    ls_sql = "Select Catcode,Description from IC_ItemCategory where compcode = '" & Gs_compcode & "' and Catcode = '" & txtdeptcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Department Code not found", vbCritical)
            Else
                txtdesc = pr_dumy("description")
               txtcatcode.SetFocus
            End If
         pr_dumy.Close
ElseIf txtdeptcode = "" And KeyCode = vbKeyReturn Then
Command1_Click
End If
End Sub

Private Sub txtsubcatcode_Change()
If txtsubcatcode = "" Then
txtsubcatdesc = ""
End If
End Sub

Private Sub txtsubcatcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtsubcatcode) <> "" And KeyCode = vbKeyReturn Then
        txtsubcatcode.Text = DoPad(txtsubcatcode.Text, 3)
        pr_dumy.Open "Select * from IC_ItemPacking where Packcode = '" & txtsubcatcode & "'  and subcode = '" & txtcatcode & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Sub Category code not found !!!", vbCritical)
            txtsubcatcode = ""
            txtsubcatdesc = ""
            txtsubcatcode.SetFocus
        Else
            txtsubcatdesc = pr_dumy("Description")
            If txtSuppliercode.Enabled Then txtSuppliercode.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtsubcatcode) = "" And KeyCode = vbKeyReturn Then
        txtsubcatcode = ""
        txtsubcatdesc = ""
        Command3_Click
End If

End Sub

Private Sub txtSuppliercode_Change()
If txtSuppliercode <> "" Then
txtSupplierdesc = ""
End If
End Sub

Private Sub txtSuppliercode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtSuppliercode) <> "" And KeyCode = vbKeyReturn Then
        txtSuppliercode.Text = DoPad(txtSuppliercode.Text, 6)
        pr_dumy.Open "Select * from IC_Supplier where Suppliercode = '" & txtSuppliercode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Supplier code not found !!!", vbCritical)
            txtSuppliercode = ""
            txtSupplierdesc = ""
            txtSuppliercode.SetFocus
        Else
            txtSupplierdesc = pr_dumy("Description")
            If DTPfrom.Enabled Then DTPfrom.SetFocus
            
        End If
        pr_dumy.Close
ElseIf Trim(txtSuppliercode) = "" And KeyCode = vbKeyReturn Then
        txtSuppliercode = ""
        txtSupplierdesc = ""
        Command4_Click
End If

End Sub

