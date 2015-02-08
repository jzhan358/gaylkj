
<!--#include file="top.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<TITLE>备份数据库</TITLE>

</HEAD>
<BODY leftmargin="20">
<table width="98%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> 　
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25"> <div align="left" style="background-image: url('../Images/topbg1.gif')"><strong>备份数据库</strong></div>
            <div align="left">
              <%
if request("action")="Backup" then
	'call backupdata()
else
%>
            </div></td>
        </tr>
        <tr> 
          <form method="post" action="backup.asp?action=Backup">
            <td> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="2" height="333">
                <tr> 
                  <td height="25" class="tdbg"><strong>备份数据文件</strong>[<a href="asp.asp">需要FSO权限</a>]</td>
                </tr>                
                <tr> 
                  <td height="22" class="tdbg">备份数据库名称[如备份目录有该文件，将覆盖，如没有，将自动创建]</td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"> <input type=text size=30 name=bkDBname value=backup.mdb></td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"><div align="center"> 
                      <input name="submit" type=submit value="确定">
                    </div></td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"><br> <br>                    
                    您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
                    注意：所有路径都是相对与程序空间根目录的相对路径(<font color="#FF0000">为了安全考虑，此功能已经关闭，请联系技术备份</font>)</td>
                </tr>
                <tr> 
                  <td height="47" class="tdbg">　</td>
                </tr>
              </table></td>
          </form>
        </tr>
      </table>
      <%
	  end if
sub backupdata() 
Dbpath=db
Dbpath=server.mappath(Dbpath) 
bkfolder="../Databackup" 
bkdbname=request.form("bkdbname") 
Set Fso=server.createobject("scripting.filesystemobject") 
if fso.fileexists(dbpath) then 
If CheckDir(bkfolder) = True Then 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
else 
MakeNewsDir bkfolder 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
end if 
response.write "备份数据库成功!" 
Else 
response.write "找不到您所需要备份的文件。" 
End if 
end sub 
'------------------检查某一目录是否存在------------------- 
Function CheckDir(FolderPath) 
folderpath=Server.MapPath(".")&"\"&folderpath 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
If fso1.FolderExists(FolderPath) then 
'存在 
CheckDir = True 
Else 
'不存在 
CheckDir = False 
End if 
Set fso1 = nothing 
End Function 
'-------------根据指定名称生成目录--------- 
Function MakeNewsDir(foldername) 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
Set f = fso1.CreateFolder(foldername) 
MakeNewsDir = True 
Set fso1 = nothing 
End Function 
%>
    </td>
  </tr>
</table>
</BODY>
</HTML>
