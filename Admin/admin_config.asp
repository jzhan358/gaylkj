﻿
<!--#include file="top.asp"-->
<!--#include file="inc/config_body.asp"-->

<title>公司基本信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
dim founderr,errmsg
founderr=false
errmsg=""
dim c_web,c_incname,c_linkman,c_addr,c_email,c_tel,c_fax
if Request.form("MM_insert") then

if request.form("action")="addconfig" then
sql="select * from config"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if  Not(rs.BOF or rs.EOF) then
   rs.close
   set rs=Nothing
   founderr=true
   errmsg=errmsg+"<br>"+"<li>已经有了基本信息，不能增加了，请使用修改功能！"
   Call diserror()
   response.End
elseif rs.BOF and rs.EOF then

rs.addnew
c_web=trim(replace(request.form("c_web"),"'",""))
c_incname=trim(replace(request.form("c_incname"),"'",""))
c_linkman=trim(replace(request.form("c_linkman"),"'",""))
c_addr=trim(replace(request.form("c_addr"),"'",""))
c_email=trim(replace(request.form("c_email"),"'",""))
c_tel=trim(replace(request.form("c_tel"),"'",""))
c_fax=trim(replace(request.form("c_fax"),"'",""))
'英文信息
e_web=trim(replace(request.form("e_web"),"'",""))
e_incname=trim(replace(request.form("e_incname"),"'",""))
e_linkman=trim(replace(request.form("e_linkman"),"'",""))
e_addr=trim(replace(request.form("e_addr"),"'",""))
e_email=trim(replace(request.form("e_email"),"'",""))
e_tel=trim(replace(request.form("e_tel"),"'",""))
e_fax=trim(replace(request.form("e_fax"),"'",""))

if c_web="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入网址！"
else
  rs("c_web")=c_web
End if

if c_incname="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入公司名称！"
else
  rs("c_incname")=c_incname
End if

if c_linkman="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入联系人！"
else
  rs("c_linkman")=c_linkman
End if
if c_addr="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入公司地址！"
else
  rs("c_addr")=c_addr
End if
if c_email="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入电子邮箱！"
else
  rs("c_email")=c_email
End if
if c_tel="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入联系电话！"
else
  rs("c_tel")=c_tel
End if
if c_fax="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入传真号码！"
else
  rs("c_fax")=c_fax
End if
rs("e_web")=e_web
rs("e_incname")=e_incname
rs("e_linkman")=e_linkman
rs("e_addr")=e_addr
rs("e_email")=e_email
rs("e_tel")=e_tel
rs("e_fax")=e_fax

if Not founderr then
rs.update
rs.close
set rs=Nothing
  closedatabase
response.redirect "admin_config.asp"
else
Call diserror()
response.End
End if
End if

End if

if request.form("action")="editconfig" then
sql="select * from config"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.BOF and rs.EOF  then
   rs.close
   set rs=Nothing
   founderr=true
   errmsg=errmsg+"<br>"+"<li>没有输入基本信息，不能修改了，请先输入基本信息！"
   Call diserror()
   response.End
elseif  Not(rs.BOF or rs.EOF) then
'dim c_web,c_incname,c_linkman,c_addr,c_email,c_tel,c_fax
c_web=trim(replace(request.form("c_web"),"'",""))
c_incname=trim(replace(request.form("c_incname"),"'",""))
c_linkman=trim(replace(request.form("c_linkman"),"'",""))
c_addr=trim(replace(request.form("c_addr"),"'",""))
c_email=trim(replace(request.form("c_email"),"'",""))
c_tel=trim(replace(request.form("c_tel"),"'",""))
c_fax=trim(replace(request.form("c_fax"),"'",""))
'英文信息
e_web=trim(replace(request.form("e_web"),"'",""))
e_incname=trim(replace(request.form("e_incname"),"'",""))
e_linkman=trim(replace(request.form("e_linkman"),"'",""))
e_addr=trim(replace(request.form("e_addr"),"'",""))
e_email=trim(replace(request.form("e_email"),"'",""))
e_tel=trim(replace(request.form("e_tel"),"'",""))
e_fax=trim(replace(request.form("e_fax"),"'",""))

if c_web="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入网址！"
else
  rs("c_web")=c_web
End if

if c_incname="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入公司名称！"
else
  rs("c_incname")=c_incname
End if

if c_linkman="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入联系人！"
else
  rs("c_linkman")=c_linkman
End if
if c_addr="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入公司地址！"
else
  rs("c_addr")=c_addr
End if
if c_email="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入电子邮箱！"
else
  rs("c_email")=c_email
End if
if c_tel="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入联系电话！"
else
  rs("c_tel")=c_tel
End if
if c_fax="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入传真号码！"
else
  rs("c_fax")=c_fax
End if
rs("e_web")=e_web
rs("e_incname")=e_incname
rs("e_linkman")=e_linkman
rs("e_addr")=e_addr
rs("e_email")=e_email
rs("e_tel")=e_tel
rs("e_fax")=e_fax

if Not founderr then

rs.update
rs.close
set rs=Nothing

response.redirect "admin_config.asp"
else
Call diserror()
response.End
End if
End if
End if
End if

Call adminconfig_body()
closedatabase

%>