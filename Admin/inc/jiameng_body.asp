﻿
<%sub jiameng_body()
dim totalnews,Currentpage,totalpages,i
sql="select * from jiameng order by id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8"><LINK 
href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td align="center" valign="middle" bgcolor="#E8E8E8">加盟管理</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">简要信息</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
        </tr>
<%if Not rs.EOF then
dim newsperpage
dim page1
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
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("id")%>　</td>
          <td bgcolor="#FFFFFF" class="chinese">[<span class="date"><%=rs("jiamengtime")%></span>]&nbsp;<a href="admin_jiameng.asp?id=<%=rs("id")%>&action=view">公司名称:<%=rs("company")%>,地址:<%=rs("addr")%>,电话:<%=rs("tel")%>,传真:<%=rs("fax")%></a> 　</td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="admin_jiameng.asp?id=<%=rs("id")%>&action=view">查看</a> <a href="admin_jiameng.asp?id=<%=rs("id")%>&action=del">删除</a></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有加盟信息！</td>
        </tr>
<%End if
End if%>
</table>
    </table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_jiameng.asp">
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
              <a href="admin_jiameng.asp?page=<%=i%>"><%=i%></a> 
              <%End if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
End if%>
              <a href="admin_jiameng.asp?page=<%=page%>" title="下一页">>></a> 
              <%End if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>

	  <%
if request.QueryString("action")="view" then
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
sql="select * from jiameng where id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if Not(rs.BOF or rs.EOF) then
%>

      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#e5e5e5">
			<form action="admin_jiameng.asp" method="post" name="form0" target="_self">
			<input name="action" type="hidden" value="del">
			<input name="id" type="hidden" value=<%= cint(request.querystring("id")) %>>
			                     <tr bgcolor="#FFFFFF">
                                    <td height="25" colspan="2" bgcolor="#F5F5F5">&nbsp;查看加盟信息&nbsp;&nbsp;&nbsp;&nbsp;</td>
              </tr>
                                  <tr bgcolor="#FFFFFF">
                                    <td width="13%" height="25" align="right" valign="middle">&nbsp;&nbsp;公司名称:</td>
                                    <td height="25" align="left" valign="middle"><%= rs("company") %></td>
                                  </tr>
                                  
                                  <tr bgcolor="#FFFFFF">
                                    <td height="25" align="right" valign="middle">&nbsp;&nbsp;联系地址:</td>
                                    <td height="25" bgcolor="#FFFFFF"><%= rs("addr") %></td>
                                  </tr>
                                  
                                  <tr bgcolor="#FFFFFF">
                                    <td height="12" align="right" valign="middle" bgcolor="#FFFFFF">&nbsp;&nbsp;联系电话:</td>
                                    <td height="12" bgcolor="#FFFFFF"><%= rs("tel") %></td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF">
                                    <td height="12" align="right" valign="middle" bgcolor="#FFFFFF">主要业务:</td>
                                    <td height="12" bgcolor="#FFFFFF"><%= rs("yewu") %></td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF">
                                    <td align="right" valign="middle">传　　真:</td>
                                    <td bgcolor="#FFFFFF"><%= rs("fax") %></td>
                                  </tr>
								  <tr bgcolor="#FFFFFF">
								    <td height="12" align="right" valign="middle" bgcolor="#FFFFFF">&nbsp;&nbsp;法人代表:</td>
                                    <td height="-2" bgcolor="#FFFFFF"><%= rs("linkman") %></td>
								  </tr>
								  <tr bgcolor="#FFFFFF">
								    <td height="12" align="right" valign="middle" bgcolor="#FFFFFF">加盟原因:</td>
		                            <td height="-1" bgcolor="#FFFFFF"><%= rs("jiamengyuanyin") %></td>
			  </tr>
								  <tr bgcolor="#FFFFFF">
								    <td height="25" align="right" valign="middle" bgcolor="#FFFFFF">加盟优势:</td>
		                            <td height="0" bgcolor="#FFFFFF"><%= rs("jiamengyoushi") %></td>
			  </tr>
								  <tr bgcolor="#FFFFFF">
								    <td height="25" align="right" valign="middle" bgcolor="#FFFFFF">加盟日期:</td>
		                            <td height="25" bgcolor="#FFFFFF"><%= rs("jiamengtime") %></td>
			  </tr>
                                  <tr align="center" valign="middle" bgcolor="#FFFFFF">
                                    <td height="30" colspan="2"><input type="submit" name="Submit" value=" 删除 " class="button">
&nbsp;&nbsp;
<input type="button" name="Submit" value=" 返回 " onClick="javascript:history.back()" class="button"></td>
                                  </tr>

	    </form>
</table>
	  <%
	  End if
	  End if
	  %>
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
End sub%>