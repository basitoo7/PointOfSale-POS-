Attribute VB_Name = "Module1"
Global Gs_RegisterTo As String
Global SetToolBar(1 To 4) As Boolean
Global Const Gs_InvldMsg    As String = "Invalid Data Found."
Global Const Gs_RecNFMsg    As String = "Record Not Found."
Global Const Gs_RecFdMsg    As String = "Record already Exists."
Global Const Gs_GlRepoPath  As String = "\IC_Reports"
Global Const Gl_Demo       As Boolean = False
Global Const gn_MaxTrns   As Integer = 150

Global Gs_compcode As String
Global Gs_CompName As String
Global Gs_Fnperiod As String
Global Gs_Txperiod As String
Global Gs_FnEndPeriod As String
Global Gs_TxEndPeriod As String
Global Gc_UserId As String
Global Gs_CPeriod As String

Global PO_AnyForm As Object
Global Para_Rs As Recordset
Private grp_rs As Recordset
Private Rights_Rs As Recordset
Global gc_dbcon As Connection
Sub Main()
    Call OpenDatabase
    Rights_Rs.Open "Select * from SysRights where UserId= '" & Gc_UserId & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
     
     If grp_rs.EOF Or grp_rs.BOF Then
        Call SetErr("Unlisenced Software Installation.", vbCritical)
        Exit Sub
     Else
        frmSplash.Show
     End If
End Sub

Private Sub OpenDatabase()
    Set gc_dbcon = New Connection
    Set Para_Rs = New Recordset
    Set Rights_Rs = New Recordset
    Set grp_rs = New Recordset
    
    gc_dbcon.CursorLocation = adUseClient
    gc_dbcon.Open "PROVIDER=MSDASQL;dsn=KimFin;uid=sa;pwd=;"
    
    Para_Rs.Open "Select * from paracount", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    grp_rs.Open "Select * from Sysregs", gc_dbcon, adOpenStatic, adLockReadOnly
    
    If Not grp_rs.EOF Then
       Gs_RegisterTo = grp_rs.Fields("GroupName")
    End If
End Sub

Public Function DoPad(Ps_String As String, Pn_Length As Integer, Optional Ps_Char As String = "0", Optional Ps_Side As String = "L") As String
    If Ps_Side = "L" Then
        DoPad = String(Pn_Length - Len(Ps_String), Ps_Char) & Ps_String
    Else
        DoPad = Ps_String & String(Pn_Length - Len(Ps_String), Ps_Char)
    End If
End Function


Public Function DentMode(RecMode, ButtNo, Tb_Rs, FrmName, codeFld As Object, NameFld, Counter_Rs, CntFld As String, CntSize, CodeStr As String, NameStr As String, Manual, Append, Toolbarx As Toolbar) As String
On Error Resume Next
      
If Range(ButtNo, 1, 3) Then
    Toolbarx.Buttons(1).Enabled = SetToolBar(1)
    Toolbarx.Buttons(2).Enabled = SetToolBar(2)
    Toolbarx.Buttons(3).Enabled = SetToolBar(3)
End If

Select Case ButtNo
  Case 1               'Addition
      Toolbarx.Buttons(1).Enabled = False
      DentMode = "A"
      If Manual = 0 Then      'If system generated counters are used
        codeFld = Format(LTrim(Str(Counter_Rs(CntFld) + 1)), Myformat(CntSize))
        codeFld.Enabled = False
        NameFld.SetFocus
        FrmName.ClearVal
      Else
        codeFld.Enabled = True
        codeFld.SetFocus
        FrmName.ClearVal
      End If
  Case 2                'Editing
        Toolbarx.Buttons(2).Enabled = False
        DentMode = "E"
        codeFld.Enabled = True
        codeFld.SetFocus
        FrmName.ClearVal
  Case 3                'Deletion
         Toolbarx.Buttons(3).Enabled = False
         DentMode = "D"
         codeFld.Enabled = True
         codeFld.SetFocus
         FrmName.ClearVal
  Case 4 And Not RecMode = ""     'Save
        
        If Not RecMode = "D" Then
            If Not FrmName.ChkInputs Then   ' Check must entered inputs
               DentMode = RecMode
               Exit Function
            End If
        End If
        
        If RecMode = "A" Then
           If Manual = 0 Then   ' if system generated counter are used saving set.
'                Para_Rs(CntFld) = Para_Rs(CntFld) + 1
'                Para_Rs.Update
           End If
           If Append Then  ' if record will be appended or not
            '  Tb_Rs.AddNew
           End If
           FrmName.SaveValues
        ElseIf RecMode = "E" Then
           FrmName.SaveValues
        ElseIf RecMode = "D" Then
           choice = SetErr("Are you sure.  ", vbYesNo)
           If choice = vbYes Then
                If Append Then
             '      Tb_Rs.Delete
                Else
                   FrmName.SaveValues
                End If
           End If
        End If

        'Tb_Rs.Update
        FrmName.ClearVal
        codeFld.Enabled = True
        codeFld.SetFocus
        DentMode = RecMode
  Case 4 And RecMode = ""
        Call SetErr("You cannot save view on disk.", vbCritical)
  Case 5                 ' Printing
        FrmName.Setprint
  Case 7                 'Cancel
        codeFld.Enabled = True
        codeFld.SetFocus
        FrmName.ClearVal
        Toolbarx.Buttons(1).Enabled = SetToolBar(1)
        Toolbarx.Buttons(2).Enabled = SetToolBar(2)
        Toolbarx.Buttons(3).Enabled = SetToolBar(3)
        DentMode = ""
End Select
End Function
Private Function Myformat(mSize) As String
    For Counter = 1 To mSize
        Myformat = Myformat + "0"
    Next
End Function
'Public Sub SetCombo(CdFld As Object, Table_Rs, Code As String, Name As String)
' If CdFld.ListCount = 0 Then
'    Table_Rs.MoveFirst
'    Do While Not Table_Rs.EOF
'       CdFld.AddItem (Table_Rs(Code) & " --> " & Table_Rs(Name))
'       Table_Rs.MoveNext
'    Loop
' End If
' Table_Rs.MoveFirst
'End Sub
'
'Private Sub UpdateCombo(CdFld As Object, Name As TextBox, LocalMode, Size)
'   If LocalMode = "E" Then
'       CdFld.AddItem (Left(CdFld, Size) & " --> " & Name)
'       CdFld.RemoveItem (CdFld.ListIndex)
'   ElseIf LocalMode = "D" Then
'       CdFld.RemoveItem (CdFld.ListIndex)
'   End If
'   CdFld.Refresh
'End Sub
'
'Public Sub SetCombo2(CdFld As ComboBox, Table_Rs, Code As String, Name As String, Name2 As String)
' If CdFld.ListCount = 0 Then
'    Table_Rs.MoveFirst
'    Do While Not Table_Rs.EOF
'       CdFld.AddItem (Table_Rs(Code) & " --> " & Table_Rs(Name) & " --> " & Table_Rs(Name2))
'       Table_Rs.MoveNext
'    Loop
' End If
' Table_Rs.MoveFirst
'End Sub


'Public Function DentShift(mDate As Date, mTime As Variant) As String
'
'        If (TimeValue(mTime) >= TimeSerial(7, 0, 1)) And (TimeValue(mTime) <= TimeSerial(15, 59, 59)) Then
'           DentShift = "A"
'        ElseIf (TimeValue(mTime) >= TimeSerial(16, 0, 0)) And (TimeValue(mTime) <= TimeSerial(23, 59, 59)) Then
'           DentShift = "B"
'        ElseIf (TimeValue(mTime) >= TimeSerial(0, 0, 1)) And (TimeValue(mTime) <= TimeSerial(0, 59, 59)) Then
'           DentShift = "C"
'           mDate = mDate - 1
'        ElseIf (TimeValue(mTime) >= TimeSerial(1, 0, 1)) And (TimeValue(mTime) <= TimeSerial(6, 59, 59)) Then
'           DentShift = "C"
'           mDate = mDate - 1
'        End If
'End Function

'Public Function SetInteger(mWeight) As Integer
'   yy = Right(LTrim(Str(mWeight)), 1)
'   SetInteger = mWeight + IIf(Val(yy) < 5, -Val(yy), 10 - Val(yy))
'End Function
Public Function MySeek(SeekValu, SeekField As String, Tabl_Rs) As Boolean
       On Error Resume Next
       Tabl_Rs.MoveFirst
       Tabl_Rs.Find (Trim(SeekField) & " = '" & SeekValu & "' ")
       MySeek = IIf(Tabl_Rs.EOF, False, True)
End Function
Public Function Range(Number, R1, R2) As Boolean
       Range = IIf(Number >= R1 And Number <= R2, True, False)
End Function

Public Function DeCode(Word, Size, Key) As String
Dim KeyChr As String
Dim Keynum As Integer

If Key = 1 Then
    KeyChr = Right(Word, 1)
    xx = Asc(KeyChr)
    Keynum = Asc(KeyChr) - 60 - (Size - 20)
Else
    Keynum = 49
End If

   For Count = 1 To Size
      DeCode = DeCode + Chr(Abs(Asc(Mid(Word, Count, 1)) - Count - Keynum - 60))
   Next
End Function

Public Function EnCode(Word, Size, Key) As String

Dim KeyChr As String
Dim Keynum As Integer

Size = IIf(Size = 0, 10, Size)
If Key = 1 Then
    KeyChr = Right(Word, 1)
    Keynum = Asc(KeyChr)
Else
    Keynum = 49
End If
   For Count = 1 To Size
      EnCode = EnCode + Chr(Asc(Mid(Word, Count, 1)) + Count + Keynum + 60)
   Next
   EnCode = EnCode + Chr(Keynum + 60 + (Size - 20))

End Function

'Public Sub LoadUser(TbRs, mBar As Object)
'Right_1 = DeCode(TbRs("Rights1"), Len(TbRs("Rights1")) - 1, 0)
'Right_2 = DeCode(TbRs("Rights2"), Len(TbRs("Rights2")) - 1, 0)
'Dim tt As Integer
'tt = 0
'   For Count = 0 To 44
'       MSetupS(Count) = IIf(Val(Mid(Right_1, Count + 1, 1)) = 0, False, True)
'       mBar.Value = tt
'       tt = tt + 1
'   Next
'
'   xx = 46
'   For Count = 0 To 49
'       mProcess(Count) = IIf(Val(Mid(Right_1, xx, 1)) = 0, False, True)
'       mReports(Count) = IIf(Val(Mid(Right_2, Count + 1, 1)) = 0, False, True)
'       mBar.Value = tt
'       xx = xx + 1
'       tt = tt + 1
'   Next
'
'   xx = 51
'   For Count = 0 To 4
'       mUtilities(Count) = IIf(Val(Mid(Right_2, xx, 1)) = 0, False, True)
'       mBar.Value = tt
'       xx = xx + 1
'       tt = tt + 1
'   Next
'End Sub

Public Function SetErr(DispMsg As String, msgtype As VbMsgBoxStyle) As Integer
     SetErr = MsgBox(DispMsg, msgtype, "Kimsys Financials.")
End Function

Public Function LastKey(ln_Key As Integer) As Boolean
     If ln_Key = vbKeyReturn Or ln_Key = vbKeyTab Then
        LastKey = True
     Else
        LastKey = False
     End If
End Function

Public Function chkRights(ProcId As String) As Boolean
  If Gc_UserId <> "ADMIN" Then
      If MySeek(ProcId, "ProcCode", Rights_Rs) Then
        chkRights = IIf(Rights_Rs.Fields("ProcRights").Value = 1, True, False)
      Else
        chkRights = False
      End If
  Else
      chkRights = True
  End If
End Function
'Private Function dummy()
'Dim ls_OrderBy As String
'Dim ls_ChActSQL As String
'Dim ls_AcctLen(0 To 9)   As String
'Dim ls_PickLen(0 To 9)   As String
'Dim ls_PlusFld(0 To 9)   As String
'Dim ln_Cnt As Integer
'Dim ln_StPos  As Integer
'Dim ln_Ucnt  As Integer
'ls_ChActSQL = "SELECT "
'
'   For ln_Cnt = 0 To gn_Maxlevels
'      ls_AcctLen(ln_Cnt) = "SubString(Gl_Detail.Acct_Sub," + LTrim(Str(ln_StPos + 1)) & "," & LTrim(Str(gn_sublen(ln_Cnt))) + ")"
'      ln_StPos = ln_StPos + gn_sublen(ln_Cnt)
'      ls_PickLen(ln_Cnt) = "left(Gl_Detail.Acct_Sub," & LTrim(Str(ln_StPos)) & ")"
'      ls_OrderBy = ls_OrderBy & "Sub" & LTrim(Str((ln_Cnt))) + ","
'
'      For ln_Ucnt = (ln_Cnt - 1) To ln_Cnt
'          If ln_Ucnt >= 0 Then
'              ls_PlusFld(ln_Cnt) = ls_PlusFld(ln_Cnt) & "Gl_sub" & LTrim(Str(ln_Cnt)) & ".Acct_Sub" & LTrim(Str(ln_Ucnt))
'              ls_PlusFld(ln_Cnt) = ls_PlusFld(ln_Cnt) & IIf(ln_Ucnt < ln_Cnt, "+", "")
'          End If
'      Next
'   Next
'   ls_OrderBy = ls_OrderBy & "Acct_Detail desc"
'
'   For ln_Ucnt = 0 To gn_Maxlevels + 1
'        For ln_Cnt = 0 To gn_Maxlevels + 1
'           If ln_Cnt = ln_Ucnt Then
'              If ln_Cnt <= gn_Maxlevels Then
'                 ls_ChActSQL = ls_ChActSQL & ls_AcctLen(ln_Ucnt) & " As Sub" & LTrim(Str(ln_Ucnt)) & ","
'              Else
'                 ls_ChActSQL = ls_ChActSQL & " Gl_Detail.Acct_Detail As Acct_Detail,"
'              End If
'           ElseIf ln_Cnt <= gn_Maxlevels Then
'               ls_ChActSQL = ls_ChActSQL & "NULL as sub" & LTrim(Str(ln_Cnt)) & ","
'           End If
'        Next
'        If ln_Ucnt <= gn_Maxlevels Then
'           ls_ChActSQL = ls_ChActSQL & "NULL as Acct_Detail,space(1) as Acct_Base,space(3) as Crncy_Code,"
'           ls_ChActSQL = ls_ChActSQL & "Gl_Sub" & LTrim(Str(ln_Ucnt)) & ".Acct_Desc As Descrip"
'           ls_ChActSQL = ls_ChActSQL & IIf(ln_Ucnt = 0, " INTO Tempdb..COA", "")
'           ls_ChActSQL = ls_ChActSQL & " From Gl_Detail,Gl_Sub" & LTrim(Str(ln_Ucnt))
'           ls_ChActSQL = ls_ChActSQL & " Where Gl_Detail.Compcode = '" & Gs_compcode & "' and " & ls_PickLen(ln_Ucnt) & "= " & ls_PlusFld(ln_Ucnt)
'           ls_ChActSQL = ls_ChActSQL & " Union All Select "
'        Else
'           ls_ChActSQL = ls_ChActSQL & " gl_Detail.acct_Base as Acct_Base,gl_Detail.Crncy_Code as Crncy_code,Acct_Desc as Descrip FROM GL_Detail "
'           ls_ChActSQL = ls_ChActSQL & " Where Gl_Detail.Compcode = '" & Gs_compcode & "' order By Sub0 DEsc "
'        End If
'   Next
'
'   gc_dbcon.Execute (ls_ChActSQL)
'
'End Function

Public Function YearClose() As Boolean
Dim PR_Syfins As New Recordset
Dim PR_Sytax As New Recordset
Dim ls_Option As Integer
Dim lb_Ok As Boolean

PR_Syfins.Open "Select * from SysFins Where compcode = '" & Gs_compcode & "' order by CompCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
PR_Sytax.Open "Select * from SysTax Where compcode = '" & Gs_compcode & "' order by CompCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1


If MySeek(Gs_Fnperiod, "ffromdate", PR_Syfins) And Format(DateValue(Date), "YYYY/MM/DD") >= Format(DateValue(Gs_FnEndPeriod), "YYYY/MM/DD") Then
   
   If SetErr("Are you sure to close year.", vbYesNo) = vbYes Then

      PR_Syfins.Fields("Fclosed") = "1"
      PR_Syfins.Fields("factiveyear") = "0"
      PR_Syfins.Update

      PR_Syfins.AddNew
      PR_Syfins.Fields("compcode") = Gs_compcode
      PR_Syfins.Fields("ffromdate") = DateValue(Gs_FnEndPeriod) + 1
      PR_Syfins.Fields("ftodate") = DateAdd("YYYY", 1, DateValue(Gs_FnEndPeriod))
      PR_Syfins.Fields("Fclosed") = "0"
      PR_Syfins.Fields("factiveyear") = "1"
      PR_Syfins.Update
      Gs_Fnperiod = PR_Syfins.Fields("ffromdate")
      Gs_FnEndPeriod = PR_Syfins.Fields("ftodate")

      If MySeek(Gs_Txperiod, "tfromdate", PR_Sytax) Then

        PR_Sytax.Fields("Tclosed") = "1"
        PR_Sytax.Fields("Tactiveyear") = "0"
        PR_Sytax.Update

        PR_Sytax.AddNew
        PR_Sytax.Fields("compcode") = Gs_compcode
        PR_Sytax.Fields("Tfromdate") = DateValue(Gs_TxEndPeriod) + 1
        PR_Sytax.Fields("Ttodate") = DateAdd("YYYY", 1, DateValue(Gs_TxEndPeriod))
        PR_Sytax.Fields("Tclosed") = "0"
        PR_Sytax.Fields("Tactiveyear") = "1"
        PR_Sytax.Update
        Gs_Txperiod = PR_Sytax.Fields("Tfromdate")
        Gs_TxEndPeriod = PR_Sytax.Fields("Ttodate")
      End If
       lb_Ok = SetCalc("Y")
       lb_Ok = IIf(lb_Ok, SetErr("Year Has successfully been closed.", vbInformation), SetErr("Error occured during year updateing opening balances", vbInformation))
   End If
End If

End Function

Public Function SetCalc(Ps_CallId As String) As Boolean
Dim cntsql As New ADODB.Command
Dim PR_GlTrans As New Recordset
Dim PR_GlRef As New Recordset
Dim ls_GlSql As String
Dim ld_stdate As Date
Dim ld_enddate As Date
SetCalc = True

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

ld_stdate = DateAdd("YYYY", -1, DateValue(Gs_Fnperiod))
ld_enddate = DateAdd("YYYY", -1, DateValue(Gs_FnEndPeriod))

ls_GlSql = "SELECT Accountno,SUM(DR_AMOUNT) AS DR_AMOUNT,SUM(CR_AMOUNT) AS CR_AMOUNT,'0OB' as VchrType,'0000000001' As Voucher_No, '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' as value_Date,'" & Gs_compcode & "' as Compcode,1 AS SerialNo,'" & Gc_UserId & "' as UserId,'" & Format(Date, "YYYY/MM/DD") & "' as AddDate ,'" & Time & "' as AddTime "
ls_GlSql = ls_GlSql + " from gl_Trans where compcode = '" & Gs_compcode & "' and value_date between '" & Format(ld_stdate, "YYYY/MM/DD") & "' AND '" & Format(ld_enddate, "YYYY/MM/DD") & "'"
ls_GlSql = ls_GlSql + " Group By AccountNo"

If Ps_CallId = "U" Then
   If SetErr("Are You Sure.", vbYesNo) = vbNo Then
      Call SetErr("Opening Balances has not been updated", vbInformation)
      Exit Function
   End If
   cntsql.CommandText = "DELETE FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' AND VchrType = '0OB' And value_date = '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "'"
   cntsql.Execute

   cntsql.CommandText = "DELETE FROM Gl_Trans WHERE CompCode = '" & Gs_compcode & "' AND VchrType = '0OB' And value_date = '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "'"
   cntsql.Execute
End If

cntsql.CommandText = "INSERT into Gl_Ref(compcode,Value_Date,Trans_Date, Voucher_No, VchrType, userid,adddate,addtime) VALUES ('" & Gs_compcode & "','" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "','" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "','0000000001','0OB','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
cntsql.Execute

cntsql.CommandText = "INSERT into Gl_Trans(AccountNo,Dr_amount,Cr_amount,vchrType,Voucher_No,Value_Date,Compcode,SerialNo,UserId,AddDate,AddTime) " & ls_GlSql
cntsql.Execute

Call SetErr("Opening Balances has successfully been updated.", vbInformation)
End Function
