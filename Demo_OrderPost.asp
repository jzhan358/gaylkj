<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="inc/safe_info.asp"-->
<% 
'if not ChkPost then  alert "error","",1
if Cstr(Session("getcode"))<>Lcase(Cstr(Trim(Request("codestr")))) or chknull(Session("getcode"),1) or chknull(Trim(Request("codestr")),1) then
	alert "验证码不正确，请检查","",1
end if
dim sql,rs
dim company,addr,tel,fax,linkman,proname,proxh,num,bz,ordertime
bookpost=checkstr(request.form("bookpost"))
company=checkstr(request.form("company"))
addr=checkstr(request.form("addr"))
tel=checkstr(request.form("tel"))
fax=checkstr(request.form("fax"))
linkman=checkstr(request.form("linkman"))
proname=checkstr(request.form("proname"))
proxh=checkstr(request.form("proxh"))
num=checkstr(request.form("num"))
bz=checkstr(request.form("bz"))
ordertime=checkstr(request.form("ordertime"))

dim founderr,errmsg
founderr=false
errmsg=""
i=0
if not chkRange(company,2,50) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Your name can not be empty!"
End if
if not chkRange(addr,4,50) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Your address can not be empty!"
End if
if not chkRange(tel,4,50) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Your phone can not be empty!"
End if
if not chkRange(proname,2,50) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Product name can not be empty!"
End if
if not chkRange(proxh,2,50) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Model can not be empty!"
End if
if not chkrequest(num) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Fill in the correct order quantity!"
End if
if founderr then alert errmsg,"",1

sql="select * from order1"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
rs("company")=company
rs("addr")=addr
rs("tel")=tel
rs("proname")=proname
rs("proxh")=proxh
rs("num")=num
rs("fax")=fax
rs("linkman")=linkman
rs("bz")=bz
rs("ordertime")=now()
rs.update
rs.close
set rs=Nothing
Session("getcode")=""
alert "You have successfully ordered the product, we will as soon as possible","",1
%>