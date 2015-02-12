<!--#include file="top.asp"-->
<!--#include file="conn.asp"-->
<%
response.Write(getorderpage())
call closeconn()
%>
