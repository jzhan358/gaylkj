﻿
<%sub m_body()
dim totalnews,Currentpage,totalpages,i
sql="select * from menu order by m_id ASC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>

<HTML>
<HEAD>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>
<style type="text/css">
A{TEXT-DECORATION: none}
A:link {COLOR: #666666; FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:visited {COLOR: #666666; FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:active {FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:hover {BORDER-BOTTOM: 1px dotted; BORDER-LEFT-WIDTH: 1px; BORDER-RIGHT-WIDTH: 1px; BORDER-TOP-WIDTH: 1px; COLOR: #ff6600; TEXT-DECORATION: none}
BODY {
FONT-SIZE: 12px;
COLOR: #666666;
FONT-FAMILY:  宋体;
background-color: #ffffff; 
background-image: url(images/show.gif);
SCROLLBAR-FACE-COLOR: #e8e7e7; 
SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
SCROLLBAR-SHADOW-COLOR: #ffffff; 
SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
SCROLLBAR-ARROW-COLOR: #ff6600; 
SCROLLBAR-TRACK-COLOR: #EFEFEF; 
SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
SCROLLBAR-BASE-COLOR: #000000
}
TABLE {BORDER-COLLAPSE: collapse; FONT-FAMILY: 宋体; FONT-SIZE: 9pt}
.button{height:18px;width:62px;background:#f6f6f9 url(images/ButtonBg.gif); border:solid 1px #5589AA;color: #000000 ;FONT-SIZE: 9pt}
.lanyu{border:solid 1px #5589AA;color: #000000 ; font-size: 12px;}
.font {  filter: DropShadow(Color=#cccccc, OffX=2, OffY=1, Positive=2); text-decoration: none; font-size: 9pt}
</style>
<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<style type="text/css">
<!--
.style1 {color: #990000}
-->
</style>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td bgcolor="#E8E8E8">主菜单管理</td>
          <td colspan="5" bgcolor="#FFFFFF"></td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="30%" bgcolor="#F5F5F5">菜单名称</td>
          <td class="chinese" width="17%" bgcolor="#F5F5F5">显示</td>
          <td class="chinese" width="18%" bgcolor="#F5F5F5">位置</td>
          <td class="chinese" width="11%" bgcolor="#F5F5F5">排序</td>
          <td class="chinese" width="14%" bgcolor="#F5F5F5">操作</td>
        </tr>
<%if Not rs.EOF then
dim newsperpage
newsperpage=10
rs.movefirst
rs.pagesize=newsperpage
if trim(request("page"))<>"" then
   page1=trim(request("page"))
   if check_num(page1)=true then
       currentpage=clng(request("page"))
       if currentpage>rs.pagecount then
          currentpage=rs.pagecount
       End if
   else
      currentpage=1
   end if
else
   currentpage=1
End if
   totalnews=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*newsperpage<totalnews then
	       rs.move(currentpage-1)*newsperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   End if
   End if
   if (totalnews mod newsperpage)=0 then
      totalpages=totalnews\newsperpage
   else
      totalpages=totalnews\newsperpage+1
   End if
i=0
Do While Not rs.EOF and i<newsperpage%>
        <tr> 
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("m_id")%></td>
          <td bgcolor="#FFFFFF" class="chinese"><%=rs("m_name")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><% if rs("m_show")=true then Response.Write("显示") else Response.Write("隐藏") End if%></td>
          <td align="center" valign="middle" bgcolor="#FFFFFF" class="chinese"><%=rs("m_position")%></td>
          <td align="center" valign="middle" bgcolor="#FFFFFF" class="chinese"><%=rs("m_sort")%></td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="Admin_Menu.asp?id=<%=rs("m_id")%>&action=editp#editp">修改</a> 
          </td>
        </tr>
<%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="6" class="chinese">当前没有分类！</td>
        </tr>
<%End if
End if%>
</table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="Admin_Menu.asp">
        <tr>
          <td class="chinese" align="right"><%=currentpage%> /<%=totalpages%>页,<%=totalnews%>条记录/<%=newsperpage%>篇每页. 
              <%
i=1
showye=totalpages
if showye>10 then
showye=10
End if
for i=1 to showye
if i=currentpage then
%>
              <%=i%> 
              <%else%>
              <a href="Admin_Menu.asp?page=<%=i%>"><%=i%></a> 
              <%End if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
End if%>
              <a href="Admin_Menu.asp?page=<%=page%>" title="下一页">>></a> 
              <%End if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>
<%
if request.QueryString("action")="editp" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>请指定操作的对象！"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>非法的ID参数！"
	Call diserror()
	response.End
  End if
End if
sql="select * from menu where m_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td colspan="4" bgcolor="#E8E8E8"><a name="editp">修改分类</a></font>（<span class="style1">说明：为了菜单的美观请把菜单的名称控制在四个汉字以内，排序的数字越小排的越前！</span>）</td>
          </tr>
          <tr align="left" valign="middle"> 
            <td width="30%" bgcolor="#F5F5F5" class="chinese">&nbsp;菜单名称</td>
            <td width="29%" bgcolor="#F5F5F5" class="chinese">&nbsp;&nbsp;&nbsp;位置</td>
            <td width="28%" bgcolor="#F5F5F5" class="chinese">&nbsp;显示</td>
            <td width="13%" bgcolor="#F5F5F5" class="chinese">&nbsp;排序</td>
          <tr align="left" valign="middle"> 
            <td bgcolor="#FFFFFF" class="chinese">
            &nbsp;<input name="mname" type="text" size="18" maxlength="18" value="<%= rs("m_name") %>"></td>
            <td bgcolor="#FFFFFF" class="chinese">&nbsp; <input type="radio" name="mp" value="顶部" <% if rs("m_position")="顶部" then Response.Write("checked")%>>
            顶部<input type="radio" name="mp" value="底部" <% if rs("m_position")="底部" then Response.Write("checked")%>>
            底部</td>
            <td bgcolor="#FFFFFF" class="chinese">&nbsp;<input name="mshow" type="radio" value="true" <% if rs("m_show")=true then Response.Write("checked")%>>显示
			<% if rs("m_name1")<>"首页" then%><input name="mshow" type="radio" value="false" <% if rs("m_show")=false then Response.Write("checked")%>>隐藏<% End if %>
			</td>
            <td bgcolor="#FFFFFF" class="chinese">&nbsp;
            <input name="msort" type="text" id="msort" size="4" maxlength="4" value="<%= rs("m_sort") %>" <% if rs("m_name1")="首页" then Response.Write("readonly='true'") End if%></td>
          <tr> 
            <td height="30" colspan="4" align="center" bgcolor="#F5F5F5" class="chinese"> 
              <input type="submit" name="Submit" value="确定修改" class="button"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="Admin_Menu.asp">返回</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("m_id")%>">
		  <input type="hidden" name="action" value="editp">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if%>
      <br>
    </td>
  </tr>
  <tr> 
    <td colspan="3" height="1" background="images/dotlineh.gif"></td>
  </tr>
</table>
<%
rs.close
set rs=Nothing
End sub
%>