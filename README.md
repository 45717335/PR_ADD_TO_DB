# PR_ADD_TO_DB
添加待上传采购申请 至数据库
## how to install
* Unzip, PR_ADD_TO_DB.7z useing the password:"PASSWORD"
* or recreate the PR_ADD_TO_DB.xlsm Useing the VBA code in scr,
## 更新
* 20200115 增加读取.txt 文件中 email信息的函数 [函数连接](https://github.com/45717335/Winshuttle_PR/blob/master/src/PR_UPLOAD.xlsm/MOD_PR_Uploading.bas)
```VBA
Private Function get_email_address(fln As String) As String
```
