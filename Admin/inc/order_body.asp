﻿
<%sub order_body()
dim totalnews,Currentpage,totalpages,i
sql="select * from order1 order by id DESC"
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
          <td align="center" valign="middle" bgcolor="#E8E8E8">订购管理</td>
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
          
    <td bgcolor="#FFFFFF" class="chinese" align="center">[<%=rs("id")%>]</td>
          
    <td bgcolor="#FFFFFF" class="chinese">&nbsp;<a href="admin_order.asp?id=<%=rs("id")%>&action=vieworder">客户名称:<font color="#FF0000"><%=rs("company")%></font>,产品名称:<font color="#FF0000"><%=rs("proname")%></font>,产品型号:<font color="#FF0000"><%=rs("proxh")%></font>,订购数量:<font color="#FF0000"><%=rs("num")%></font></a> 
      　</td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="admin_order.asp?id=<%=rs("id")%>&action=vieworder">查看</a> <a href="admin_order.asp?id=<%=rs("id")%>&action=delorder">删除</a></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          
    <td bgcolor="#FFFFFF" colspan="3" class="chinese"><font color="#FF0000"><strong>暂无定购产品信息！</strong></font></td>
        </tr>
<%End if
End if%>
</table>
    </table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_order.asp">
        <tr>
          
      <td class="chinese" align="right"> <%if currentpage>1 then%><a href="admin_order.asp?page=<%=currentpage-1%>" title="上一页">上一页</a><%end if%>&nbsp;<%if currentpage<totalpages then%><a href="admin_order.asp?page=<%=currentpage+1%>" title="下一页">下一页</a><%end if%>&nbsp;当前页/总共页:<font color=red><%=currentpage%></font>/<font color=red><%=totalpages%></font>,当前记录/每页总共记录 <font color=red><%=i%></font> /<font color=red><%=newsperpage%></font> 
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>

	  <%
if request.QueryString("action")="vieworder" then
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
sql="select * from order1 where id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if Not(rs.BOF or rs.EOF) then
%>

      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#e5e5e5">
			<form action="admin_order.asp" method="post" name="form0" target="_self">
			<input name="action" type="hidden" value="delorder">
			<input name="id" type="hidden" value=<%= cint(request.querystring("id")) %>>
			                     <tr bgcolor="#FFFFFF">
                                    
      <td height="25" colspan="6" bgcolor="#F5F5F5">&nbsp;查看订购信息&nbsp;&nbsp;&nbsp;&nbsp;客户名称：<font color="#FF0000"><%=rs("company")%></font>&nbsp;&nbsp;订购时间:<font color="#FF0000"><%=rs("ordertime")%></font></td>
              </tr>
								  <%proname=rs("proname")
								  proxh=rs("proxh")
								  num=rs("num")
								  proname = Split(proname,",")
								  proxh = Split(proxh,",")
								  num = Split(num,",")
								  for i=0 to ubound(proname)%>
                                  <tr bgcolor="#FFFFFF">
                                    <td width="13%" height="25" align="right" valign="middle">产品名称：</td>
                                    
      <td width="16%" height="25">　<%=rs("proname")%></td>
                                    <td height="25">产品型号：</td>
                                    
      <td height="25">　<%=rs("proxh")%></td>
                                    <td height="25">订购数量：</td>
                                    
      <td height="25">　<%=rs("num")%></td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF">
                                    <td height="25" align="right" valign="middle">&nbsp;&nbsp;联系地址:</td>
                                    
      <td height="25" colspan="5">　<%=rs("addr")%></td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF">
                                    <td align="right" valign="middle">联系电话:</td>
                                    
      <td bgcolor="#FFFFFF">　<%=rs("tel")%></td>
                                    <td width="9%" bgcolor="#FFFFFF">传　　真:</td>
                                    
      <td width="20%" bgcolor="#FFFFFF">　<%=rs("fax")%></td>
                                    <td width="11%" bgcolor="#FFFFFF">联系人:</td>
                                    
      <td width="31%" bgcolor="#FFFFFF">　<%=rs("linkman")%></td>
                                  </tr>
								  <tr bgcolor="#FFFFFF">
								    <td height="25" align="right" valign="middle">&nbsp;&nbsp;备注信息:</td>
                                    
      <td height="25" colspan="5">　<%=rs("bz")%></td><%next%>
								  </tr>
                                  <tr align="center" valign="middle" bgcolor="#FFFFFF">
                                    <td height="30" colspan="6"><input type="submit" name="Submit" value=" 删除 " class="button">
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