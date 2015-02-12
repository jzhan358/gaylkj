<!--#include file="top.asp"-->
<!--#include file="conn.asp"-->
<%
sql="select news_id,news_language from news_class"
tarr=getmdbvaluelist(sql)
for i=0 to ubound(tarr,2)
	sql="update news set news_language="&tarr(1,i)&" where news_class_id="&tarr(0,i)
	call setmdbvalue(sql)
next
call closeconn()
%>
