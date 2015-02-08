
<!--#include file="top.asp"-->
<!--#include file="inc/order_body.asp"-->
<title>订购管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
dim founderr,errmsg
founderr=false
errmsg=""

if request.form("action")="delorder" or request.querystring("action")="delorder" then
if request.Form("id")="" AND request.querystring("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) AND not isInteger(request.querystring("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的订购id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
'删除订购信息
if request.form("id")="" then
   sql="select * from order1 where id="&cint(request.querystring("id"))
else
   sql="select * from order1 where id="&cint(request.Form("id"))
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.delete
rs.close
set rs=Nothing
  closedatabase
Response.write "<script language = 'javascript'>alert('成功删除此订购信息!');window.document.location.href='admin_order.asp';</script>"
response.end
'End if
End if

Call order_body()
closedatabase
End if
%>