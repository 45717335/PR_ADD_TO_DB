# PR_ADD_TO_DB
添加待上传采购申请 至数据库
## how to install
* Unzip, PR_ADD_TO_DB.7z useing the password:"PASSWORD"
* or recreate the PR_ADD_TO_DB.xlsm Useing the VBA code in scr,
## 更新
* 20200311 自动填写Username
* 20200115 增加mm函数收集每个用户的输入（用户名.txt 里面 存放邮件） [函数连接](https://github.com/45717335/PR_ADD_TO_DB/blob/master/src/PR_ADD_TO_DB.xlsm/Module1.bas)
```VBA
Private Function get_email_address(fln As String) As String
Private Function write_file(fln As String, s_text As String, s_path As String) As Boolean
```
![图片](https://github.com/45717335/PR_ADD_TO_DB/blob/master/PIC/%E6%94%B6%E9%9B%86%E9%82%AE%E4%BB%B6%E5%9C%B0%E5%9D%80.png)
