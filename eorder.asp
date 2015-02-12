<!--#include file="top.asp"-->
<!--#include file="conn.asp"-->
<%
response.Write(geteorderpage())
call closeconn()
%>