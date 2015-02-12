<!--#include file="top.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="inc/page.asp"-->
<%
response.Write(getnewspage())
call closeconn()
%>
