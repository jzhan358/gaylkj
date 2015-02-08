<!--#include file="../top.asp"-->
<%dns="../"%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/md5.asp"-->
<% 
dim aname,apass,FoundErr,ErrMsg
FoundErr=False
aname=checkstr(request.Form("name"))
apass=checkstr(request.Form("pass"))
safecode=checkstr(Request.Form("safecode"))
if not chkRange(aname,3,20)  then
   FoundErr=True
   ErrMsg=ErrMsg&"用户名不对！\n\n"
End if

if not chkRange(apass,6,20)  then
   FoundErr=True
   ErrMsg=ErrMsg&"用户密码不对！\n\n"
End if
if safeCode<>cstr(Session("Admin_GetCode")) or chkNull(safecode,1) then
	FoundErr=True
	ErrMsg=ErrMsg & "验证码不正确！\n\n"
end if

if FoundErr then alert ErrMsg,"",1


apass=md5(apass)
dim sql,rs
sql="select a_name,a_pass,a_flag from admin where a_name='"&aname&"' and a_pass='"&apass&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.BOF and rs.EOF then
	rs.close
	set rs=Nothing
    ErrMsg="用户名或密码错误！"
	alert ErrMsg,"",1
end if
session("aname")=rs("a_name")
session("admin_flag")="into"
session("admin_sys")=rs("a_flag")
rs.close
set rs=Nothing
conn.close
set conn=Nothing
response.Redirect("./")
 %>
