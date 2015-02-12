<!--#include file="top.asp"-->
<!--#include file="conn.asp"-->
<% 
id=request.QueryString("id")
if not chkrequest(id) then alert "error","",1
response.Write(getproduct_viewpage(id))
call closeconn()
%>