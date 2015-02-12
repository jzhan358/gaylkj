<!--#include file="top.asp"-->
<!--#include file="conn.asp"-->
<%
t=request.QueryString("type")
id=request.QueryString("id")
if t="1" then t=1 else t=0
't=0 为新闻。目前只有新闻
if not chkrequest(id) then 
response.Write("NaN")
response.End()
end if
sql="select news_count from news where news_id="&id
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
	closers(rs)
	response.Write("NaN")
	response.End()
end if
rs("news_count")=rs("news_count")+1
rs.update
response.Write(rs("news_count"))
closers(rs)
closeconn()
response.End()
%>
