<!--#include file="top.asp"-->
<%
dim founderr,errmsg
founderr=false
errmsg=""
dim p_type
p_type=checkstr(request.form("p_type"))
p_type_e=checkstr(request.form("p_type_e"))
id=request.Form("id")
if action="new" or action="edit" then
	if not chkrange(p_type,1,20) then alert "请输入中文栏目名称(1-20字)","",1
	if not chkrange(p_type_e,1,20) then alert "请输入英文栏目名称(1-20字符)","",1
end if
if action="new" or action="del" then 
	if not chkrequest(id) then alert "error","",1
end if
action=request.Form("action")
if action="new" then		
	sql="select * from p_class"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew	
	rs("p_type")=p_type	
	rs("p_type_e")=p_type_e	
	rs.update
	rs.close
	set rs=Nothing	
End if
if action="edit" then	
	sql="select * from p_class where p_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3	
	rs("p_type")=p_type	
	rs("p_type_e")=p_type_e	
	rs.update
	rs.close	
	set rs=Nothing	
	sql="update p_class_small set p_type='"&p_type&"',p_type_e='"&p_type_e&"' where p_type_id="&id
	conn.execute(sql)
	sql="update p_info set p_type='"&p_type&"',p_type_e='"&p_type_e&"' where p_type_id="&id
	conn.execute(sql)	
End if

if action="del" then
	sql="delete * from p_class_small where p_type_id="&id
	conn.execute(sql)
	sql="delete * from p_info where p_type_id="&id
	conn.execute(sql)			
	sql="select * from p_class where p_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.delete
	rs.close
	set rs=Nothing	
End if

closeconn
response.redirect "proclass.asp"
'End if
%>