VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGenRecipereport 
   Caption         =   "Generate Production Stock"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenRecipe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1500
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
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
      Height          =   1605
      Left            =   30
      TabIndex        =   5
      Top             =   -120
      Width           =   4275
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   3150
         TabIndex        =   4
         Top             =   1200
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
         Left            =   2070
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   1200
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   4035
         Top             =   210
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
      Begin VB.TextBox txtVchrDesc 
         Height          =   315
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   1065
         Left            =   75
         TabIndex        =   6
         Top             =   90
         Width           =   4125
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   2175
            TabIndex        =   1
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   54460417
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   2175
            TabIndex        =   2
            Top             =   630
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   54460417
            CurrentDate     =   37293
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   240
            Left            =   105
            TabIndex        =   10
            Top             =   735
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   1410
            TabIndex        =   8
            Top             =   630
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1230
            TabIndex        =   7
            Top             =   255
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frmGenRecipereport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Dim pr_dumy1 As New Recordset
Dim PR_Dumy2 As New Recordset

Public codeid As String
Dim ls_sql As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
If Label1.Caption = "Gen" Then
        Dim ls_ItemCode As String
        Dim ls_RecipeCode As String
        Dim ln_qty As Double
        Dim ln_Amount As Double
        Dim ln_RecipeRate As Double
        Dim res
        
        ls_sql = "delete from IC_RecipesStock where compcode = '" & Gs_compcode & "' and TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' and TransDate <= '" & Format(DTPTo, "YYYY/MM/DD") & "' "
        gc_dbcon.Execute ls_sql
        
        
        ls_sql = "SELECT  IC_TransAutoSale.ItemCode, sum(IC_TransAutoSale.Quantity) as Quantity FROM IC_TransAutoSale INNER JOIN IC_TransMasterAutoSale ON IC_TransAutoSale.Compcode = IC_TransMasterAutoSale.Compcode AND  IC_TransAutoSale.TransCode = IC_TransMasterAutoSale.TransCode where IC_TransMasterAutoSale.compcode = '" & Gs_compcode & "' and IC_TransMasterAutoSale.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "'and IC_TransMasterAutoSale.TransDate <= '" & Format(DTPTo, "YYYY/MM/DD") & "'  group by IC_TransAutoSale.Itemcode  order by IC_TransAutoSale.Itemcode "
        pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If Not pr_dumy.EOF Then
          Do While Not pr_dumy.EOF
            ls_ItemCode = Trim(pr_dumy("ItemCode") & "")
            If Trim(ls_ItemCode) <> "" Then
            ls_sql = "SELECT  * from IC_RecipeFormula where compcode = '" & Gs_compcode & "' and ItemCode = '" & ls_ItemCode & "' order by recipecode"
            pr_dumy1.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If Not pr_dumy1.EOF Then
                
                Do While Not pr_dumy1.EOF
                
                ln_qty = Val(0 & pr_dumy("Quantity")) * Val(0 & pr_dumy1("Quantity"))
                ls_RecipeCode = Trim(pr_dumy1("RecipeCode") & "")
                
                ls_sql = "SELECT * from IC_Recipes where recipecode = '" & ls_RecipeCode & "'"
                
                PR_Dumy2.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
                If Not PR_Dumy2.EOF Then
                    ln_RecipeRate = Val(0 & PR_Dumy2("RecipeRate"))
                End If
                PR_Dumy2.Close
                
                
                ln_Amount = ln_qty * ln_RecipeRate
                
                ls_sql = "Insert into IC_RecipesStock(Compcode, TransDate, ItemCode, RecipeCode, Quantity, RecipeRate, Amount)"
                ls_sql = ls_sql & " Values ('" & Gs_compcode & "','" & Format(Date, "YYYY/MM/DD") & "','" & ls_ItemCode & "','" & ls_RecipeCode & "'," & ln_qty & "," & ln_RecipeRate & "," & ln_Amount & ")"
                gc_dbcon.Execute ls_sql
                
                pr_dumy1.MoveNext
                Loop
                
            Else
                Call MsgBox("Formula not set for itemcode " + ls_ItemCode, vbCritical)
                res = MsgBox("do you want to continuse Process !!!", vbYesNo + vbCritical)
                If res = vbNo Then
                 pr_dumy1.Close
                 pr_dumy.Close
                 Exit Sub
                End If
            End If
            pr_dumy1.Close
            End If
            
          pr_dumy.MoveNext
          Loop
          
        Else
            Call MsgBox("Stock not found !!!", vbCritical)
        End If
        pr_dumy.Close
        
        Call MsgBox("Successfully Generated", vbInformation)
        
        With crrpt
                .ReportFileName = App.Path & Gs_ICRepoPath & "\RecipeStock.rpt"
                .WindowTitle = Me.Caption
                .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
                .Formulas(1) = "Reportname = 'Recipes Stock Ledger'"
                .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
                .SelectionFormula = "{IC_RecipesStock.CompCode} = '" & Gs_compcode & "'"
                .SelectionFormula = .SelectionFormula & " AND {IC_RecipesStock.Transdate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {IC_RecipesStock.Transdate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"
                .Connect = "DNS=Censoft;UID=Sa"
                .Action = 1
        End With
Else
        With crrpt
                .ReportFileName = App.Path & Gs_ICRepoPath & "\RecipeStock.rpt"
                .WindowTitle = Me.Caption
                .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
                .Formulas(1) = "Reportname = 'Recipes Stock Ledger'"
                .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
                .SelectionFormula = "{IC_RecipesStock.CompCode} = '" & Gs_compcode & "'"
                .SelectionFormula = .SelectionFormula & " AND {IC_RecipesStock.Transdate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {IC_RecipesStock.Transdate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"
                .Connect = "DNS=Censoft;UID=Sa"
                .Action = 1
        End With

End If

Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Form_Load()
  
  dtpfrom = Date
  DTPTo = Date

End Sub
