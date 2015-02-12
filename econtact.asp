<!--#include file="top.asp"-->
<!--#include file="conn.asp"-->
<%
response.Write(getecontactpage())
call closeconn()
%>