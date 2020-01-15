# PR_ADD_TO_DB
添加待上传采购申请 至数据库
## how to install
* Unzip, PR_ADD_TO_DB.7z useing the password:"PASSWORD"
* or recreate the PR_ADD_TO_DB.xlsm Useing the VBA code in scr,
## 更新
* 20200115 增加mm函数收集每个用户的输入（用户名.txt 里面 存放邮件） [函数连接](https://github.com/45717335/Winshuttle_PR/blob/master/src/PR_UPLOAD.xlsm/MOD_PR_Uploading.bas)
```VBA
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
```
