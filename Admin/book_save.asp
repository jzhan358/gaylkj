<!--#include file="top.asp"-->
<%
dim founderr,errmsg
founderr=false
errmsg=""
action=request("action")
id=request("id")
if action="del" then
	if not chkrequest(id) then alert "error","",1
	sql="select * from msg where m_id="&id
	set rs=server.createobject("adodb.recordset")	
	rs.open sql,conn,1,3
	if not (rs.bof and rs.eof) then
	rs.delete
	else
	alert "记录已经不存在了","",1
	end if
	rs.close
	set rs=Nothing
End if

closeconn
response.Redirect("book.asp")
%>