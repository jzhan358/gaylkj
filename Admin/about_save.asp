<!--#include file="top.asp"-->
<%
dim founderr,errmsg
founderr=false
errmsg=""

'if Request.form("MM_insert") then
Language=request.Form("Language")
if Language<>"1" then Language="0"
if request.form("action")="newabout" then
	dim i_type
	i_type=checkstr(request.form("i_type"))
	i_language=request.form("i_language")
	if not chkrange(i_type,1,20) then alert "请输入栏目名称(1-20字)","",1
	if i_language<>"0" and i_language<>"1" then alert "没有选择语言","",1
	
	sql="select * from inc_class"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew	
	rs("i_type")=i_type	
	rs("i_language")=i_language	
	rs.update
	rs.close
	set rs=Nothing		
End if
if request.form("action")="editabout" then
	id=request.Form("id")
	i_type=checkstr(request.form("i_type"))
	i_language=request.form("i_language")
	if not chkrequest(id) then alert "error","",1
	if not chkrange(i_type,1,20) then alert "请输入栏目名称(1-20字)","",1
	if i_language<>"0" and i_language<>"1" then alert "没有选择语言","",1
	sql="select * from inc_class where i_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3	
	rs("i_type")=i_type	
	rs("i_language")=i_language	
	rs.update
	rs.close	
	set rs=Nothing	
	
End if

if request.form("action")="delabout" then
	id=request.Form("id")
	if not chkrequest(id) then alert "error","",1		
	sql="select * from inc_class where i_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.delete
	rs.close
	set rs=Nothing
	
	v_sql="select * from inc_info where inc_class_id="&id
	set v_rs=server.createobject("adodb.recordset")
	v_rs.open v_sql,conn,1,3
	if not v_rs.eof then
	   v_rs.delete
	end if
	v_rs.close
	set v_rs=Nothing	
End if


if request.form("action")="edittextabout" then
	id=request.Form("id")
	inc_text=request.form("inc_text")
	if not chkrequest(id) then alert "error","",1	
	if chkNull(inc_text,10) then alert "内容必须填写(10字符以上)","",1	
	sql="select * from inc_info where inc_class_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs("inc_text")=replace_text(inc_text)
	rs.update
	rs.close
	set rs=Nothing	
	'生成静态
	if cityviewmode=1 then
		select case id
		case "14"'中文的公司介绍
			tarr=savetohtml(getintropage(),"/intro.html")
			tarr2=savetohtml(getindexpage(),"/index.html")
		case "25"'中文的联系我们
			tarr=savetohtml(getcontactpage(),"/contact.html")
			tarr2=savetohtml(getindexpage(),"/index.html")
		case "34"'英文的公司介绍
			tarr=savetohtml(geteintropage(),"/eintro.html")
			tarr2=savetohtml(getdefaultpage(),"/default.html")
		case "35"'英文的联系我们		
			tarr=savetohtml(getecontactpage(),"/econtact.html")
			tarr2=savetohtml(getdefaultpage(),"/default.html")
		end select
		if tarr(0)<>0 then alert tarr(1),"",1
		if tarr2(0)<>0 then alert tarr2(1),"",1
		erase tarr	
		call createSeoPage()			
	end if	
End if

if request.form("action")="addabout" then
	id=request.Form("id")
	inc_text=request.form("inc_text")
	if not chkrequest(id) then alert "error","",1	
	if chkNull(inc_text,10) then alert "内容必须填写(10字符以上)","",1	
	sql="select * from inc_info where inc_class_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if rs.EOF and rs.BOF then
		rs.addnew
		rs("inc_text")=replace_text(inc_text)
		rs("inc_class_id")=id
		rs.update
	End if
	rs.close
	set rs=Nothing	
End if
closeconn
response.redirect "about.asp?language="&language
'End if
%>