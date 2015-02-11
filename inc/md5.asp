<!--#include file="md52.asp"-->

<%
function md5(byVal str)
md5=md52(left(md52(str),17))
End Function
%>