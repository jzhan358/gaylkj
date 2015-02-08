
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
          <td align="center" valign="middle" bgcolor="#E8E8E8">订购管理</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">描叙</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
        </tr>
<%
dim sql_id,sql_Field,sql_table,sql_order,sqlst
dim pagesize,pagecount,recordcount,page
'条件
sqlst="id>0"&st
'统计字段   
sql_id="id"  
'显示字段   
sql_Field="id,company,proname,proxh,num"  
'查询表名   
sql_table="order1"  
'排序字段   
sql_order="id"  
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
          <td bgcolor="#FFFFFF" class="chinese" align="center">[<%=rs("id")%>]　</td>
          <td bgcolor="#FFFFFF" class="chinese">&nbsp;<a href="order.asp?id=<%=rs("id")%>&action=view">客户名称:<font color="#FF0000"><%=rs("company")%></font>,产品名称:<font color="#FF0000"><%=rs("proname")%></font>,产品型号:<font color="#FF0000"><%=rs("proxh")%></font>,订购数量:<font color="#FF0000"><%=rs("num")%></font></font></td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="order.asp?id=<%=rs("id")%>&action=view">查看</a> <a href="order_save.asp?id=<%=rs("id")%>&action=del">删除</a></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有订购！</td>
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
sql="select * from order1 where id="&id
rs.open sql,conn,1,1
if Not(rs.BOF or rs.EOF) then
%>

     <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#e5e5e5">
			<form action="order_save.asp" method="post" name="form0">
			<input name="action" type="hidden" value="del">
			<input name="id" type="hidden" value=<%= id %>>
			                     <tr bgcolor="#FFFFFF">
                                    
      <td height="25" colspan="6" bgcolor="#F5F5F5">&nbsp;查看订购信息&nbsp;&nbsp;&nbsp;&nbsp;客户名称：<font color="#FF0000"><%=rs("company")%></font>&nbsp;&nbsp;订购时间:<font color="#FF0000"><%=rs("ordertime")%></font></td>
              </tr>		
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
                                    
      <td height="25" colspan="5">　<%=rs("bz")%></td>
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
end if

set rs=nothing
call closeconn()
%>