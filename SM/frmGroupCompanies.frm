VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGroupCompanies 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group Companies"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGroupCompanies.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdGroupCompanies 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      MousePointer    =   14
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Select"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectCompany 
         Caption         =   "Select Company"
      End
   End
End
Attribute VB_Name = "frmGroupCompanies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PR_SysComp As Recordset
Dim PR_SyRights As New Recordset

Private Sub Form_Load()
MDIForm1.StatusBar1.Panels(7).Text = ""
PR_SyRights.Open "Select syrights.*,Ltrim(Rtrim(syrights.processid)) as Findfld from syrights  where userid = '" & Gc_UserId & "' Order by processid", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    InitializeGrid
    FillGrid
 
End Sub

Public Sub InitializeGrid()
    With grdGroupCompanies
        .Redraw = False
        .Rows = 2
        .FormatString = "Code|<Description|<Printingname"
        .ColWidth(0) = 500
        .ColWidth(1) = 4100
        .ColWidth(2) = 0
        .Redraw = True
    End With
End Sub

Private Sub FillGrid()
On Error GoTo LocalErr
    Set PR_SysComp = New Recordset
    
    PR_SysComp.Open "SELECT SysComp.*, RTRIM(CompCode) AS Code FROM SysComp where status =1", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    
    If PR_SysComp.RecordCount = 0 Then Exit Sub
    With grdGroupCompanies
        While Not PR_SysComp.EOF
          '  If UCase(Gc_UserId) <> "ADMIN" Then
           '     If MySeek(Trim(PR_SysComp!Code), "Findfld", PR_SyRights) Then
            '        .Row = .Rows - 1
             '       .TextMatrix(.Row, 0) = PR_SysComp!Code
              '      .TextMatrix(.Row, 1) = PR_SysComp!CompName
               '     .TextMatrix(.Row, 2) = PR_SysComp!compcash
                    
                '    .Rows = .Rows + 1
               ' End If
            'Else
                    .Row = .Rows - 1
                    .TextMatrix(.Row, 0) = PR_SysComp!Code
                    .TextMatrix(.Row, 1) = PR_SysComp!CompName
                    .TextMatrix(.Row, 2) = PR_SysComp!compcash
                    
                    .Rows = .Rows + 1
           ' End If
            PR_SysComp.MoveNext
        Wend
        .Row = 1
         If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
    End With
     
    Exit Sub
LocalErr:
End Sub

Public Sub SetValues(Optional ls_compcode As String)
    With grdGroupCompanies
        If .Rows = 2 Then
            If Val(.TextMatrix(1, 0)) = 0 Then
                Call SetErr("No Company available for selection.", vbInformation)
                Exit Sub
            End If
        End If
    End With
    
    Dim ln_cnt     As Integer
    Dim PR_SysFins As Recordset
    Dim PR_SysTax  As Recordset
    
    Gn_TotLen = 0
    Set PR_SysFins = New Recordset
    Set PR_SysTax = New Recordset
    
    PR_SysFins.Open "SELECT SysFins.*, RTRIM(CompCode)+factiveyear AS Code FROM SysFins", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    PR_SysTax.Open "SELECT SysTax.*, RTRIM(CompCode)+tactiveyear AS Code FROM SysTax", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    
    With grdGroupCompanies
        Gs_compcode = .TextMatrix(.Row, 0)
        Gs_CompName = .TextMatrix(.Row, 2)
        Gs_SelectedCompName = .TextMatrix(.Row, 1)
        
        PR_SysComp.MoveFirst
        PR_SysComp.Find ("Code = '" & Trim(Gs_compcode) & "'")
        If Not PR_SysComp.EOF Then
            gn_Maxlevels = IIf(Gb_GL, Val(0 & PR_SysComp!GLActLevel), 0)
            gn_DtlLen = IIf(Gb_GL, Val(0 & PR_SysComp.Fields("glActdetl")), 0)
            Gs_BranchCode = PR_SysComp.Fields("BranchCode") & ""
            Gs_BaseCrncy = PR_SysComp!BaseCurrency & ""
            gs_ICBase = IIf(Gb_IC, PR_SysComp.Fields("Inv_Base") & "", "")
            gs_ArBase = IIf(Gb_AR, PR_SysComp.Fields("AR_Base") & "", "")
            Para_Rs.Filter = "Compcode = '" & Gs_compcode & "'"
            GoTop Para_Rs
            If Gb_GL Then
                For ln_cnt = 0 To 9
                    gn_sublen(ln_cnt) = Val(0 & PR_SysComp.Fields("glactsub" + LTrim(str(ln_cnt))))
                    Gn_TotLen = Gn_TotLen + gn_sublen(ln_cnt)
                Next
                Gn_TotLen = Gn_TotLen + Val(0 & PR_SysComp.Fields("glactDetl"))
            End If
        Else
           Call SetErr("company Not Found.", vbCritical)
           Unload Me
           Exit Sub
        End If
                
        If PR_SysFins.RecordCount > 0 Then
            PR_SysFins.MoveFirst
            PR_SysFins.Find ("Code = '" & (Trim(Gs_compcode) & "1") & "'"), 0, adSearchForward, 1
        End If
        If Not PR_SysFins.EOF Then
            Gs_Fnperiod = PR_SysFins!FFromDate
            Gs_FnEndPeriod = PR_SysFins!FToDate
            If ParaCntr_Rs.State = 1 Then ParaCntr_Rs.Close
            ParaCntr_Rs.Open "Select SysFins.* From SysFins Where Compcode = '" & Gs_compcode & "' And fActiveYear = '1'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
            GoTop ParaCntr_Rs
        End If
        
        If PR_SysTax.RecordCount > 0 Then
            PR_SysTax.MoveFirst
            PR_SysTax.Find ("Code = '" & (Trim(Gs_compcode) & "1") & "'"), 0, adSearchForward, 1
            Gs_SetPeriod = "Value_Date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' And Value_Date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "'"
        End If
        
        If Not PR_SysTax.EOF Then
            Gs_Txperiod = PR_SysTax!TFromDate
            Gs_TxEndPeriod = PR_SysTax!TToDate
        End If
    End With
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With grdGroupCompanies
        If .Rows = 2 Then
            If Val(.TextMatrix(.Row, 0)) = 0 Then
                Unload Me
                'Unload MDISt
            End If
        End If
        If Not Val(.TextMatrix(1, 0)) = 0 And Trim(Gs_compcode) = "" Then
            Call SetErr("You must select a Company.", vbInformation)
            'Cancel = True
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
PR_SyRights.Close
End Sub

Public Sub grdGroupCompanies_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SetValues
End Sub

Private Sub grdGroupCompanies_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then SetValues
End Sub

Private Sub mnuSelectCompany_Click()
    SetValues
End Sub
