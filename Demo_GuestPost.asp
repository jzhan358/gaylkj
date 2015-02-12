<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="inc/safe_info.asp"-->
<% 
if Cstr(Session("getcode"))<>Lcase(Cstr(Trim(Request("codestr")))) or chknull(Session("getcode"),1) or chknull(Trim(Request("codestr")),1) then
	alert "验证码不正确，请检查","",1
end if

'if not ChkPost then  alert "error","",1
dim sql,rs
dim m_title,m_name,m_email,m_tel,m_addr,m_fax,m_text,m_city,bookpost
bookpost=checkstr(request.form("bookpost"))
if bookpost<>"bookpost" then alert "error","",1
m_title=checkstr(request.form("m_title"))
m_name=checkstr(request.form("m_name"))
m_email=checkstr(request.form("m_email"))
m_tel=checkstr(request.form("m_tel"))
m_addr=checkstr(request.form("m_addr"))
m_fax=checkstr(request.form("m_fax"))
m_text=replace_text(request.form("m_text"))

dim founderr,errmsg
founderr=false
errmsg=""
i=0
if not chkRange(m_title,2,50) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Your title can not be empty!"
End if
if not chkRange(m_name,2,20) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Your name can not be empty!"
End if
if not chkRange(m_text,5,300) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Please enter your comments or some suggestions!"
End if
if not IsValidEmail(m_email) then
  founderr=true
  i=i+1
  errmsg=errmsg&i&")、Please enter your valid e-mail!"
End if
if founderr then alert errmsg,"",1

sql="select * from msg"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
rs("m_title")=m_title
rs("m_name")=m_name
rs("m_text")=m_text
rs("m_email")=m_email
rs("m_tel")=m_tel
rs("m_addr")=m_addr
rs("m_fax")=m_fax
rs("m_city")=m_city
rs("m_date")=now()
rs.update
rs.close
set rs=Nothing
Session("getcode")=""
alert "You have successfully message, thank you for your message, we will as soon as possible!","",1
%>
