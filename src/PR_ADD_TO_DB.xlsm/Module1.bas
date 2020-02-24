Attribute VB_Name = "Module1"
Option Explicit


#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As LongLong)
#Else
    Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If


Sub F_TOBEUL()
'已经传好的 复制的 PR_DONE
'选择文件添加值数据库 待传。
'已经传好的 复制的 PR_DONE
Dim re As String, ct As String, sql As String

re = InStr(1, Application.OperatingSystem, "64-bit", vbTextCompare)
If re = 0 Then
    ct = "provider=Microsoft.jet.OLEDB.4.0;data source="
Else
    ct = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="
End If
Dim i_last As Long
Dim j_last As Long
Dim j As Long
Dim kk as long            
Dim i As Long
Dim temp_s As String
Dim fln As String
Dim filetoopen As Variant
Application.DefaultFilePath = ThisWorkbook.Worksheets(1).Range("A1")
filetoopen = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", Title:="Choose file:", MultiSelect:=True)
If Not IsArray(filetoopen) Then
Exit Sub
Else
j = 1
'Workbooks("PR_ADD_TO_DB.xlsm").Worksheets(1).UsedRange.Clear
ThisWorkbook.Worksheets(1).UsedRange.Clear

For i = LBound(filetoopen) To UBound(filetoopen)
temp_s = filetoopen(i)
fln = Right(temp_s, Len(temp_s) - InStrRev(temp_s, "\"))
If fln Like "P?????_CN*.xlsm" Or fln Like "M?????_CN*.xlsm" Or fln Like "D?????_CN*.xlsm" Then

If ready_to_upload(temp_s) = True Then
Cells(j, 1) = temp_s
Cells(j, 2) = fln
j = j + 1
End If
Else
MsgBox "File Name Must like :  P?????_CN*.xlsm or M?????_CN*.xlsm or D?????_CN*.xlsm"
End If
Next
End If
Application.DefaultFilePath = ThisWorkbook.Worksheets(1).Range("A1") = Left(temp_s, InStr(temp_s, fln))
Dim myCon      As New ADODB.Connection
Dim myRst      As New ADODB.Recordset
Dim myRst2     As New ADODB.Recordset
Dim myFileName As String
Dim myTblName  As String
Dim myTblName2  As String
Dim myKey      As String
Dim mySht      As Worksheet
Dim str1 As String
Dim str4    As String
Dim ExcelApp As Excel.Application
Dim ExcelWB As Excel.Workbook
Set ExcelApp = GetObject(, "Excel.application")
Set ExcelWB = Nothing
myFileName = "HTML_Data.mdb"
myTblName = "PR_TOBEUPLOAD"
myTblName2 = "PR_DONE"

If re = 0 Then
'=======================================================================================================
myCon.Open ct & "Z:\24_Temp\PA_Logs\HTML\mdb\" & myFileName & ";"
myCon.Execute "SELECT * FROM " & myTblName
myRst2.Index = "PrimaryKey"
myRst2.Open Source:=myTblName2, ActiveConnection:=myCon, _
                                    CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
                                    Options:=adCmdTableDirect
'=======================================================================================================
Else
'=======================================================================================================

    myCon.Open ct & "Z:\24_Temp\PA_Logs\HTML\mdb\HTML_Data.mdb"

sql = "select * from " & myTblName2

    myRst2.Open sql, myCon, adOpenKeyset, adLockOptimistic
    
    
    
'=======================================================================================================
End If

With myRst

If re = 0 Then
'==========================================================
.Index = "PrimaryKey"
myRst.Open Source:=myTblName, ActiveConnection:=myCon, _
CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
Options:=adCmdTableDirect

'==========================================================
Else
'==========================================================
sql = "select * from " & myTblName
    myRst.Open sql, myCon, adOpenKeyset, adLockOptimistic
    
'==========================================================
End If

myRst2.MoveLast
j_last = myRst2.Fields(0).Value
i_last = 1
Do While Not .EOF
str1 = .Fields(3).Value
If .Fields(0).Value > i_last Then i_last = .Fields(0).Value
If .Fields(1).Value <> "TOBEDONE" And Left(.Fields(1).Value, 4) <> "DOIN" Then
myRst2.AddNew
j_last = j_last + 1
myRst2.Fields(0) = j_last
For i = 1 To myRst2.Fields.Count - 1
myRst2.Fields(.Fields(i).Name).Value = .Fields(i).Value
Next i
myRst2.Update
myRst.Delete
DoEvents
.Update
End If
.MoveNext
Loop
j = j - 1
For i = 1 To j
If Len(Cells(i, 1)) > 10 Then
str1 = Cells(i, 1)
Set myRst2 = myCon.Execute("SELECT * FROM PR_TOBEUPLOAD WHERE FLFP='" & str1 & "'")
If myRst2.BOF Then
.AddNew
i_last = i_last + 1
.Fields(0).Value = i_last
.Fields(1).Value = "TOBEDONE"
.Fields(3).Value = Cells(i, 1)
.Update
Else
.Seek myRst2.Fields(0).Value
If Not .BOF Then
.Fields(1).Value = "TOBEDONE"
.Update
End If
End If
End If
Next
.Close
End With
myCon.Close
Set myRst = Nothing
Set myCon = Nothing

ThisWorkbook.Saved = True
'ThisWorkbook.Application.DisplayAlerts = False
'ThisWorkbook.Close

End Sub


Sub Macro1()
'select *  from tablename  order by filename1 desc

ThisWorkbook.ActiveSheet.UsedRange.Clear
Dim re As String, ct As String, sql As String
 Dim cnn As New ADODB.Connection
 Dim rst As New ADODB.Recordset
'---------------------------------------------------------------- 判断系统是否为64位
re = InStr(1, Application.OperatingSystem, "64-bit", vbTextCompare)
If re = 0 Then
    ct = "provider=Microsoft.jet.OLEDB.4.0;data source="
Else
    ct = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="
End If
'----------------------------------------------------------------- 载入基本配置
cnn.Open ct & "Z:\24_Temp\PA_Logs\HTML\mdb\HTML_Data.mdb"
'sql = "select * from PR_TOBEUPLOAD order by 编号 desc"
sql = "select * from PR_TOBEUPLOAD"
rst.Open sql, cnn, adOpenKeyset, adLockOptimistic
Range("A2").CopyFromRecordset rst
Range("A2").Select
End Sub


Sub Macro2()
ThisWorkbook.ActiveSheet.UsedRange.Clear
Dim re As String
 Dim ct As String
 Dim sql As String
 
 Dim cnn As New ADODB.Connection
 Dim rst As New ADODB.Recordset
'---------------------------------------------------------------- 判断系统是否为64位
re = InStr(1, Application.OperatingSystem, "64-bit", vbTextCompare)
If re = 0 Then
    ct = "provider=Microsoft.jet.OLEDB.4.0;data source="
Else
    ct = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="
End If
'----------------------------------------------------------------- 载入基本配置
cnn.Open ct & "Z:\24_Temp\PA_Logs\HTML\mdb\HTML_Data.mdb"
'sql = "select top 100 * from (select * from PR_DONE order by 编号 desc)"
sql = ThisWorkbook.Worksheets("setting").Range("A1")
If InStr(sql, "select top 100 * from (select * from PR_DONE order by") = 0 Then
MsgBox "SQL Error!"
Exit Sub
End If
rst.Open sql, cnn, adOpenKeyset, adLockOptimistic
Range("A2").CopyFromRecordset rst
Range("A2").Select
End Sub

Function ready_to_upload(flfp As String) As Boolean


Application.AutomationSecurity = msoAutomationSecurityForceDisable


Application.ScreenUpdating = False
'如果所选PR单已经存在单号，报错并返回假
Dim str1 As String
Dim str2 As String, str3 As String
Dim date_formate_ddmmyyyy As Boolean
date_formate_ddmmyyyy = False
Dim i As Integer, i_last As Integer
Dim i_start As Integer
Dim ws_pr As Worksheet

Dim b_c As Boolean

Dim i_20 As Integer
Dim C_B As String, C_C As String, C_D As String, C_E As String, C_F As String, C_G As String, C_H As String, C_I As String, C_J As String, C_K As String, C_L As String, C_M As String, C_N As String, C_O As String, C_P As String, C_Q As String, C_R As String


ready_to_upload = True
'On Error GoTo Errorhand
Dim wb As Workbook

'Set wb = Workbooks.Open(flfp, 0, 1)
Set wb = Workbooks.Open(Filename:=flfp, WriteResPassword:="TKSY")


Set ws_pr = wb.ActiveSheet


If wb.ActiveSheet.Range("B10") = "Protocol:" And wb.ActiveSheet.Range("C10") = "" Then


'类型设定
i_20 = 21
C_B = "B": C_C = "C"
C_E = "D": C_G = "F"
C_I = "H": C_K = "J"
C_M = "L": C_Q = "M"
C_J = "I"
'类型设定

        '检查单号内外一致
        str3 = wb.ActiveSheet.Range("O7")
        str2 = Left(get_fln(flfp), Len(str3))
        If str3 <> str2 Then
        MsgBox str3 & "<>" & str2 & "please check before uploading!"
        add_comm str3, wb.ActiveSheet, 7, 15, True
        wb.ActiveSheet.Range("O7").Interior.Color = RGB(255, 0, 0)
        wb.ActiveSheet.Range("O7") = Left(get_fln(flfp), 6)
        ready_to_upload = False
        End If
        '检查单号内外一致
        

            str1 = wb.ActiveSheet.Range("N21")
            
            If Not (str1 Like "??.??.????") Then
            MsgBox str1 & Chr(10) & "Delivery Date must like DD.MM.YYYY"
            date_formate_ddmmyyyy = True
            ready_to_upload = False
            End If
            
            'PX????.??? 不能大于10个字符
             
            If ready_to_upload = True Then
            str1 = wb.ActiveSheet.Range("C21")
            
            If Len(str1) > 10 Then
            MsgBox str1 & Chr(10) & "PX????.???  Must < 10 Char"
            ready_to_upload = False
            End If
            
            End If
            
            
                         
                         
            If ready_to_upload = True Then
            str1 = wb.ActiveSheet.Range("D21")
            
            If Len(str1) = 0 Then
            MsgBox "Enter ShortText  D column"
            ready_to_upload = False
            End If
            
            End If
            
            
            

ElseIf wb.ActiveSheet.Range("C10") = "Protocol:" And wb.ActiveSheet.Range("D10") = "" Then


'类型设定
i_20 = 20
C_B = "B": C_C = "C": C_D = "D": C_E = "E": C_F = "F": C_G = "G": C_H = "H": C_I = "I": C_J = "J": C_K = "K": C_L = "L": C_M = "M": C_N = "N": C_O = "O": C_P = "P": C_Q = "Q": C_R = "R"
'类型设定



        '检查单号内外一致
        str3 = wb.ActiveSheet.Range("P7")
        str2 = Left(get_fln(flfp), Len(str3))
        If str3 <> str2 Then
        MsgBox str3 & "<>" & str2 & "please check before uploading!"
        add_comm str3, wb.ActiveSheet, 7, 16, True
        wb.ActiveSheet.Range("P7").Interior.Color = RGB(255, 0, 0)
        wb.ActiveSheet.Range("P7") = Left(get_fln(flfp), 6)
        ready_to_upload = False
        End If
        '检查单号内外一致
        
        
        
        
            
            str1 = wb.ActiveSheet.Range("O20")
            If Not (str1 Like "??.??.????") Then
            MsgBox str1 & Chr(10) & "Delivery Date must like DD.MM.YYYY"
            date_formate_ddmmyyyy = True
            ready_to_upload = False
            End If
            
            Else
            MsgBox "Clear Protocol before uploading!"
            ready_to_upload = False
            End If
            
            
            
            'PX????.??? 不能大于10个字符
             
            If ready_to_upload = True Then
            str1 = wb.ActiveSheet.Range("C21")
            
            If Len(str1) > 10 Then
            MsgBox str1 & Chr(10) & "PX????.???  Must < 10 Char"
            ready_to_upload = False
            End If
            
            End If
            
            If ready_to_upload = True Then
            str1 = wb.ActiveSheet.Range("E20")
            
            If Len(str1) = 0 Then
            MsgBox "Enter ShortText  E column"
            ready_to_upload = False
            End If
            End If
            




If ready_to_upload Then
b_c = True
i_last = i_20
Do While b_c = True
'E,G,J 连续两行为空则终止
If Len(Trim(ws_pr.Range(C_E & i_last + 2))) + Len(Trim(ws_pr.Range(C_G & i_last + 2))) + Len(Trim(ws_pr.Range(C_J & i_last + 2))) = 0 Then
If Len(Trim(ws_pr.Range(C_E & i_last + 1))) + Len(Trim(ws_pr.Range(C_G & i_last + 1))) + Len(Trim(ws_pr.Range(C_J & i_last + 1))) = 0 Then b_c = False
End If
If Len(Trim(ws_pr.Range(C_E & i_last))) + Len(Trim(ws_pr.Range(C_G & i_last))) + Len(Trim(ws_pr.Range(C_J & i_last))) = 0 Then
ws_pr.Rows(i_last).Clear
ws_pr.Rows(i_last).Interior.Color = RGB(255, 0, 0)
ready_to_upload = False

Else
'C_B 列
str1 = Trim(ws_pr.Range(C_B & i_last))
If Len(str1) = 0 Then
ws_pr.Range(C_B & i_last).Interior.Color = RGB(255, 0, 0)
ws_pr.Range(C_B & i_last) = i_last - i_20 + 1
ready_to_upload = False
End If
'C_C 列
str1 = Trim(ws_pr.Range(C_C & i_last))
If Len(str1) = 0 Then
ws_pr.Range(C_C & i_last).Interior.Color = RGB(255, 0, 0)
ws_pr.Range(C_C & i_last) = Left(get_fln(flfp), 6) & "." & Right("000" & i_last - i_20 + 1, 3)
ready_to_upload = False
End If
'C_J,C_G,C_Q,C_K
str1 = C_J
If Len(Trim(ws_pr.Range(str1 & i_last))) = 0 Then
ws_pr.Range(str1 & i_last).Interior.Color = RGB(255, 0, 0)
ready_to_upload = False
End If
str1 = C_G
If Len(Trim(ws_pr.Range(str1 & i_last))) = 0 Then
ws_pr.Range(str1 & i_last).Interior.Color = RGB(255, 0, 0)
ready_to_upload = False
End If
str1 = C_Q
If Len(Trim(ws_pr.Range(str1 & i_last))) = 0 Then
ws_pr.Range(str1 & i_last).Interior.Color = RGB(255, 0, 0)
ready_to_upload = False
End If
str1 = C_K
If Len(Trim(ws_pr.Range(str1 & i_last))) = 0 Then
ws_pr.Range(str1 & i_last).Interior.Color = RGB(255, 0, 0)
ready_to_upload = False
End If
'C_J,C_G,C_Q,C_K
End If
i_last = i_last + 1
Loop
End If



            
            
            
If ready_to_upload = True Then
'设置打印区域

If ws_pr.PageSetup.PrintArea <> "" Then
ws_pr.PageSetup.PrintArea = "$C$1:$P$" & Right(ws_pr.PageSetup.PrintArea, Len(ws_pr.PageSetup.PrintArea) - 8)
End If


'设置打印区域
wb.Save
wb.Close 0
Else

'设置打印区域
If ws_pr.PageSetup.PrintArea <> "" Then
ws_pr.PageSetup.PrintArea = "$B$1:$Q$" & Right(ws_pr.PageSetup.PrintArea, Len(ws_pr.PageSetup.PrintArea) - 8)
End If

'设置打印区域


End If
Application.AutomationSecurity = msoAutomationSecurityLow


Application.ScreenUpdating = True
If ready_to_upload = False Then
'修改日期格式
If date_formate_ddmmyyyy Then
If change_date_format(flfp) Then
ready_to_upload = ready_to_upload(flfp)
End If
End If
'修改日期格式
End If




Exit Function
Errorhand:
ready_to_upload = False
MsgBox Err.Description
Exit Function
End Function

Sub macrotest()
MsgBox ready_to_upload("Z:\24_Temp\PA_Logs\V1.2\PR_UPLOADED\AFT\20181222\PE0450_CN.505809_MINGXUAN_20181222.xlsm")
End Sub
Sub print_prtest()
Workbooks.Open Filename:="Z:\24_Temp\PA_Logs\TOOLS\PRINT_PR\Print_PR.xlsm", ReadOnly:=True
End Sub
Sub info()
Shell "explorer.exe " & Range("B2"), vbNormalFocus
End Sub
Private Function get_fln(flfp As String) As String
If InStr(flfp, "\") > 0 Then
get_fln = Right(flfp, Len(flfp) - InStrRev(flfp, "\"))
Else
get_fln = flfp
End If
End Function

Private Function change_date_format(flfp As String) As Boolean
'将格式转换为"DD.MM.YYYY"
Dim wb As Workbook
Dim start_i As Integer, i_last As Integer
Dim i As Integer

Dim de_da As String
Dim ws As Worksheet

Application.AutomationSecurity = msoAutomationSecurityForceDisable
Set wb = Workbooks.Open(Filename:=flfp, WriteResPassword:="TKSY")

If wb.ActiveSheet.Range("B10") = "Protocol:" And wb.ActiveSheet.Range("C10") = "" Then
start_i = 21
de_da = "N"
ElseIf wb.ActiveSheet.Range("C10") = "Protocol:" And wb.ActiveSheet.Range("D10") = "" Then
start_i = 20
de_da = "O"
End If

Set ws = wb.ActiveSheet
i_last = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
For i = start_i To i_last
ws.Range(de_da & i) = format_date_DDMMYYYY(ws.Range(de_da & i))
Next
wb.Save
wb.Close
Application.AutomationSecurity = msoAutomationSecurityLow
change_date_format = True
End Function

Sub ma()
 change_date_format "D:\temp\PE0040_CN.505887_3250AK22_20190710.xlsm"
 
End Sub


Function format_date_DDMMYYYY(m_c As Range) As String
'格式化日期函数
'支持Excel全部日期格式和CW？？形式
    Dim date_1 As Date
    Dim s_1 As String
    Dim wk As Integer
    Dim str_date As String
'===============================
'单元格已经是日期格式的，进行格式转换

    If IsDate(m_c) = True Then
    date_1 = m_c
    format_date_DDMMYYYY = Format(date_1, "DD.MM.YYYY")
    Else
    format_date_DDMMYYYY = Trim(m_c)
    End If
'单元格已经是日期格式的，进行格式转换
'===============================
'===========================
'判断是否转换成功，如果未成功，判断是否为CW##格式并转换

If format_date_DDMMYYYY Like "##.##.####" Then
'成功直接跳过
Else

    str_date = format_date_DDMMYYYY
    If str_date Like "CW?/????" Then
    str_date = "CW0" & Right(str_date, 6)
    ElseIf str_date Like "CW?/????*" Then
    str_date = "CW0" & Mid(str_date, 3, 6)
    End If
    
    If str_date Like "CW??*" Then
    'Return the sunday of special week
    wk = CInt(Mid(str_date, 3, 2))
    Dim InputNum As Integer, FirstD As Date, StartD As Date, i As Integer
    InputNum = Val(wk)
    FirstD = CDate(Year(Date) & "-1" & "-1")
    StartD = FirstD + (InputNum - 1) * 7 - Weekday(FirstD, vbMonday) + 1
    date_1 = CDate(StartD + 4)
    If date_1 < Now() Then
    If str_date Like "CW??*" Then
    'Return the sunday of special week
    wk = CInt(Mid(str_date, 3, 2))
    InputNum = Val(wk)
    FirstD = CDate((Year(Date) + 1) & "-1" & "-1")
    StartD = FirstD + (InputNum - 1) * 7 - Weekday(FirstD, vbMonday) + 1
    date_1 = CDate(StartD + 4)
    ElseIf str_date Like "????-*-*" Then
    'Return the Change directly
    date_1 = CDate(str_date)
    End If
    End If
    format_date_DDMMYYYY = Format(date_1, "DD.MM.YYYY")
    End If
    
End If
'判断是否转换成功，如果未成功，判断是否为CW##格式并转换
'===========================


'If m_c.Comment Is Nothing Then m_c.AddComment
'm_c.Comment.Text Text:=CStr(m_c)
'm_c.NumberFormat = "yyyy/mm/dd;@"
'm_c = date_1






End Function


Function add_comm(ByVal comm_s As String, ws1 As Worksheet, ByVal h_i As Integer, ByVal l_i As Integer, ByVal visiable As Boolean) As Boolean
On Error GoTo Errorhand
If ws1.Cells(h_i, l_i).Comment Is Nothing Then
    ws1.Cells(h_i, l_i).AddComment
End If
ws1.Cells(h_i, l_i).Comment.Text Text:=comm_s
ws1.Cells(h_i, l_i).Comment.Visible = visiable
Exit Function
Errorhand:
If Err.Number <> 0 Then MsgBox "get_ws function: " + Err.Description
Err.Clear
End Function

Sub mm()
Dim str1 As String, str2 As String, str3 As String
str3 = "Z:\24_Temp\PA_Logs\TOOLS\Winshuttle_auto\Email\"
'=========================Applicant:
 str1 = Application.UserName
If Len(str1) > 12 Then str1 = Environ("username")
If Len(str1) > 12 Then str1 = Left(str1, 12)
'=========================Applicant:
str2 = get_email_address(str1)
str2 = InputBox("please input your email address." & Chr(10) & str1, "PRU", str2)
If str2 Like "*@thyssenkrupp.com" Then
write_file str1, str2, str3
End If
End Sub

Private Function get_email_address(fln As String) As String
Dim str3 As String, str2 As String
Dim fs As Object, a As Object
str3 = "Z:\24_Temp\PA_Logs\TOOLS\Winshuttle_auto\Email\"
If Right(fln, 4) <> ".txt" Then fln = fln & ".txt"
Set fs = CreateObject("Scripting.FileSystemObject")
On Error GoTo Errorhand
Set a = fs.OpenTextFile(str3 & fln)
get_email_address = a.Readall
a.Close
Exit Function
Errorhand:
get_email_address = ""
End Function

Private Function write_file(fln As String, s_text As String, s_path As String) As Boolean
write_file = False
Dim str1 As String, str2 As String, str3 As String
Dim atf As Object, FSO As Object
If Right(fln, 4) <> ".txt" Then fln = fln & ".txt"
   fln = Replace(Replace(fln, "?", ""), "*", "")
        If Len(Trim(s_text)) > 0 Then
            Set FSO = CreateObject("Scripting.FileSystemObject")
            If FSO.folderexists(s_path) = False Then
                FSO.CreateFolder s_path
            End If
            If FSO.FileExists(s_path & fln) Then
            Kill s_path & fln
            End If
                Set atf = FSO.CreateTextFile(s_path & fln, True)
                atf.WriteLine (s_text)
                atf.Close
                write_file = True
        End If
End Function

