<!--#include file="top.asp"-->
<%
dim founderr,errmsg
founderr=false
errmsg=""

'if Request.form("MM_insert") then
Language=request.Form("Language")
if Language<>"1" then Language="0"
action=request.Form("action")
if action="new" then
	dim news_type
	news_type=checkstr(request.form("news_type"))
	news_language=request.form("news_language")
	if not chkrange(news_type,1,20) then alert "请输入栏目名称(1-20字)","",1
	if news_language<>"0" and news_language<>"1" then alert "没有选择语言","",1
	
	sql="select * from news_class"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew	
	rs("news_type")=news_type	
	rs("news_language")=news_language	
	rs.update
	rs.close
	set rs=Nothing	
End if
if action="edit" then
	id=request.Form("id")
	news_type=checkstr(request.form("news_type"))
	news_language=request.form("news_language")
	if not chkrequest(id) then alert "error","",1
	if not chkrange(news_type,1,20) then alert "请输入栏目名称(1-20字)","",1
	if news_language<>"0" and news_language<>"1" then alert "没有选择语言","",1
	sql="select * from news_class where news_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3	
	rs("news_type")=news_type	
	rs("news_language")=news_language	
	rs.update
	rs.close	
	set rs=Nothing	
End if

if action="del" then
	id=request.Form("id")
	if not chkrequest(id) then alert "error","",1		
	sql="select * from news_class where news_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.delete
	rs.close
	set rs=Nothing
	sql="delete * from news where news_class_id="&id
	conn.execute(sql)
End if

closeconn
response.redirect "newsclass.asp?language="&language
'End if
%>