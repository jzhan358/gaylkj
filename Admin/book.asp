
<!--#include file="top.asp"-->
<!--#include file="../inc/page.asp"-->
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8"><LINK 
href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td align="center" valign="middle" bgcolor="#E8E8E8">留言管理</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">标题</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
        </tr>
<%
dim sql_id,sql_Field,sql_table,sql_order,sqlst
dim pagesize,pagecount,recordcount,page
'条件
sqlst="m_id>0"&st
'统计字段   
sql_id="m_id"  
'显示字段   
sql_Field="m_id,m_title,m_date"  
'查询表名   
sql_table="msg"  
'排序字段   
sql_order="m_id"  
'每页记录   
PageSize=20   
'获得总数  
sql="select count("&sql_id&") from "&sql_table&" where "&sqlst&" "
recordcount=conn.execute(sql)(0)
	'总页数  
	if cint(recordcount) = 0 then 
		pagecount=1
	else
		pagecount=Abs(Int(recordcount/PageSize*(-1)))   
	end if
	'获得当前页码   
	page=request.QueryString("page")
	if not chkrequest(page) then page=1 else page=cint(page)
	if page>pagecount then page=pagecount

'sql语句   
if page=1 then   
	sql="SELECT TOP "&PageSize&" "&sql_Field&" from "&sql_table&" where "&sqlst&" order by "&sql_order&" desc"  
else   
	sql="SELECT TOP "&PageSize&" "&sql_Field&" from "&sql_table&" where ("&sql_order&" <(SELECT MIN("&sql_order&") FROM (SELECT TOP "&((Page-1)*PageSize)&" "&sql_order&" FROM "&sql_table&" where "&sqlst&" order by "&sql_order&" desc) AS tblTMP)) and "&sqlst&" order by "&sql_order &" desc" 
end if 

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if Not (rs.EOF and rs.eof) then
i=0
Do While Not rs.EOF and i<PageSize
%>
        <tr> 
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("m_id")%>　</td>
          <td bgcolor="#FFFFFF" class="chinese">[<span class="date"><%=rs("m_date")%></span>]&nbsp;<a href="book.asp?id=<%=rs("m_id")%>&action=view"><%=rs("m_title")%></a> 　</td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="book.asp?id=<%=rs("m_id")%>&action=view">查看</a> <a href="book_save.asp?id=<%=rs("m_id")%>&action=del">删除</a></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有留言！</td>
        </tr>
<%
End if%>
</table> 
<table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">	   
  <tr>
    <td class="chinese" align="right"><%=showpage(pagecount,pagesize,page,recordcount,15)%></td>
  </tr>	
</table>
<%rs.close%>

<%
action=request.QueryString("action")
if action="view" then
id=request.QueryString("id")
if not chkrequest(id) then alert "error","",1
sql="select * from msg where m_id="&id
rs.open sql,conn,1,1
if Not(rs.BOF or rs.EOF) then
%>

      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#e5e5e5">
			<form action="book_save.asp" method="post" name="form0">
			<input name="action" type="hidden" value="del">
			<input name="id" type="hidden" value=<%=id%>>
			                     <tr bgcolor="#FFFFFF">
                                    <td height="25" colspan="6" bgcolor="#F5F5F5">&nbsp;查看留言&nbsp;&nbsp;&nbsp;&nbsp;留言人：<%= rs("m_name") %>&nbsp;&nbsp;留言时间:<%= rs("m_date") %></td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF">
                                    <td width="12%" height="25" align="right" valign="middle">&nbsp;&nbsp;主　　题:</td>
                                    <td height="25" colspan="5" align="left" valign="middle"><%= rs("m_title") %></td>
                                  </tr>
                                  
                                  <tr bgcolor="#FFFFFF">
                                    <td height="25" align="right" valign="middle">&nbsp;&nbsp;邮件地址:</td>
                                    <td width="27%" height="25"><%= rs("m_email") %></td>
                                    <td width="11%">联系电话：</td>
                                    <td width="20%"><%= rs("m_tel") %></td>
                                    <td width="10%">传真号码：</td>
                                    <td width="20%"><%= rs("m_fax") %></td>
                                  </tr>
                                  
                                  <tr bgcolor="#FFFFFF">
                                    <td height="25" align="right" valign="middle">&nbsp;&nbsp;详细地址:</td>
                                    <td height="25" colspan="5"><%= rs("m_addr") %></td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF">
                                    <td align="right" valign="middle">留言内容:</td>
                                    <td colspan="5"><%= rs("m_text") %></td>
                                  </tr>
                                  <tr align="center" valign="middle" bgcolor="#FFFFFF">
                                    <td height="30" colspan="6"><input type="submit" name="Submit" value=" 删除 " class="button">
&nbsp;&nbsp;
<input type="button" name="Submit" value=" 返回 " onClick="javascript:history.back()" class="button"></td>
                                  </tr>

	    </form>
</table>
<%
end if
rs.close
End if


set rs=Nothing
call closeconn()
%>