<% 
border="#CECECE"
 %>
<%sub head() %>

<table width="780" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="37" align="center" background="images/bg_t1.gif"><img src="images/t001.gif" width="73" height="37" border="0" usemap="#Map2"><img src="images/t002.gif" width="105" height="37" border="0" usemap="#Map"><img src="images/t003.gif" width="105" height="37" border="0" usemap="#Map3"><img src="images/t004.gif" width="105" height="37" border="0" usemap="#Map4"><img src="images/t005.gif" width="105" height="37" border="0" usemap="#Map5"><img src="images/t006.gif" width="105" height="37" border="0" usemap="#Map6"><img src="images/t007.gif" width="105" height="37" border="0" usemap="#Map7"></td>
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="12,8,99,30" href="about.asp" target="_self" alt="关于我们">
</map>
<map name="Map2">
  <area shape="rect" coords="11,9,68,32" href="index.asp" target="_self" alt="首页">
</map>
<map name="Map3">
  <area shape="rect" coords="11,9,101,30" href="products/products.asp" target="_self" alt="产品展厅">
</map>
<map name="Map4">
  <area shape="rect" coords="13,10,98,31" href="news.asp" target="_self" alt="新闻速递">
</map>
<map name="Map5">
  <area shape="rect" coords="9,9,98,30" href="server.asp" target="_self" alt="客户中心">
</map>
<map name="Map6">
  <area shape="rect" coords="10,10,97,31" href="hr.asp" target="_self" alt="人才加盟">
</map>
<map name="Map7">
  <area shape="rect" coords="12,10,98,31" href="link.asp" target="_self" alt="联系我们">
</map>
<% End sub %>
<% sub footer() %>
<table border="0" cellpadding="0" cellspacing="0" width="780">
  <tr>
    <td valign="middle" height="47" background="images/fon_bot.gif"><p align="center" class="style3">Copyright &copy;2003 CompanyName.com</p></td>
    <td background="images/fon_bot.gif" valign="middle"><p class="menu02"> <a href="">站点声明</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="#">合作商机</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="#">客户服务</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="#">联系方式</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#">问题回答</a></p></td>
  </tr>
  <tr>
    <td colspan="2" height="15" bgcolor="#F77100"><img src="images/px1.gif" width="1" height="1" alt="" border="0"></td>
  </tr>
</table>
<% End sub %>
<% sub news001()
dim n_sql,n_rs 
n_sql="select top 5 news_id,news_title from news order by news_id DESC"
set n_rs=server.createobject("adodb.recordset")
n_rs.open n_sql,conn,1,1
if n_rs.BOF and n_rs.EOF then
%>
<tr>
  <td height="25">&nbsp;<img src="images/news_icon.gif" width="12" height="12">暂无信息</td>
</tr>

<%
n_rs.close
set n_rs=Nothing
elseif Not(n_rs.BOF or n_rs.EOF) then
Do While Not n_rs.EOF
%>
<tr>
  <td height="25">&nbsp;<img src="images/news_icon.gif" width="12" height="12">&nbsp;<a href="shownews.asp?news_id=<%=n_rs("news_id")%>" target="_blank"><%=left(n_rs("news_title"),20)&"..."%></a></td>
</tr>
<% 
n_rs.movenext
loop
End if
End sub
 %>
 <%sub br%>
<table border="0"><tr><td height="5"></td></tr></table>
<%End sub%>
