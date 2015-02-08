<!--#include file="top.asp"-->
<%
dim founderr,errmsg
dim author,ahome,keyword,title,news_class_id,content,news_language
founderr=false
errmsg=""
i=0
action=request.Form("action")
language=request.QueryString("language")
page=request.QueryString("page")
keyword=request.QueryString("keyword")
cid=request.QueryString("cid")
id=request.Form("id")

if action="new" or action="edit" then
	author=checkstr(request.form("news_author"))
	ahome=checkstr(request.form("news_ahome"))
	keyword=checkstr(request.form("news_keyword"))
	title=checkstr(request.form("news_title"))
	news_class_id=checkstr(request.form("news_class_id"))
	content=request.form("news_content")	
	
	if not chkRange(title,2,50) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你必须输入信息的标题！"	
	End if
	
	if not chkrequest(news_class_id) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你必须选择信息的分类！"	
	else
		sql="select news_language from news_class where news_id="&news_class_id
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			founderr=true
			i=i+1
			errmsg=errmsg&i&"),你选择信息分类并不存在！"
		else
			news_language=rs("news_language")
		end if	
		rs.close
		set rs=nothing
	End if
	if chknull(content,10) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你必须输入信息的内容（10字符以上）！"	
	End if
end if

if action="edit" or action="del" then
	if not chkrequest(id) then alert "error","",1
end if

if founderr then alert errmsg,"",1
if not chkrequest(news_language) then news_language=0 else news_language=cint(news_language)
if action="new" then
	sql="select * from news"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
	rs("news_author")=author
	rs("news_ahome")=ahome
	rs("news_keyword")=keyword
	rs("news_title")=title
	rs("news_class_id")=news_class_id
	rs("news_language")=news_language
	rs("news_content")=replace_text(content)
	rs("news_date")=date
	rs.update
	id=rs("news_id")
	rs.close
	set rs=Nothing
End if

if action="edit" then
sql="select * from news where news_id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
	rs("news_author")=author
	rs("news_ahome")=ahome
	rs("news_keyword")=keyword
	rs("news_title")=title
	rs("news_class_id")=news_class_id
	rs("news_language")=news_language
	rs("news_content")=replace_text(content)	
	rs.update
	rs.close
	set rs=Nothing
End if

if (action="new" or action="edit") and cityviewmode=1 then
	if news_language=0 then
		tarr=savetohtml(getnews_viewpage(id),"/news/"&id&".html")
		tarr2=savetohtml(getindexpage(),"/index.html")
		call createSeoPage()
	else
		tarr=savetohtml(getenews_viewpage(id),"/news/"&id&".html")
		tarr2=savetohtml(getdefaultpage(),"/default.html")
	end if
	
	if tarr(0)<>0 then alert tarr(1),"",1
end if

if action="del" then
	sql="select * from news where news_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.delete
	rs.close
	set rs=Nothing 
	call delfile("/news/"&id&".html")
	if cityviewmode=1 then
		call savetohtml(getindexpage(),"/index.html")
		call savetohtml(getdefaultpage(),"/default.html")
	end if
	call createSeoPage()
End if
closeconn
response.Redirect("news.asp?language="&language&"&page="&page&"&cid="&cid&"&keyword="&server.URLEncode(keyword))
%>