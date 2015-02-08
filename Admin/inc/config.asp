<!--#include file="../conn.asp"-->
<%
dim sql_config,rs_config
sql_config="select * from config"
set rs_config=server.createobject("adodb.recordset")
rs_config.open sql_config,conn,1,1
%>