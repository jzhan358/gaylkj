﻿
<%sub admin_case_body()
dim totalnews,Currentpage,totalpages,i
sql="select * from case_info order by p_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8"><LINK 
href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<SCRIPT language=javascript>
  var upfile_obj;
  function OnUpFile()
  {
   upfile_obj.src=upfile_obj.lowsrc;
   upfile_obj.src=document.forms["form1"].upfile.value;
  }
</SCRIPT>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td bgcolor="#E8E8E8" colspan="3">案例管理(图片上传时请注意图片的大小和格式，只可以上传大小在200K以下的“.jpg”或“.gif”结尾的图片。)</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">案例名称</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
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
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("p_id")%>　</td>
          <td bgcolor="#FFFFFF" class="chinese"><a href="../showcase.asp?p_id=<%=rs("p_id")%>" target="_blank"><%=rs("p_name")%></a></span>　</td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="admin_case.asp?id=<%=rs("p_id")%>&action=editpro#editpro">修改</a> 
            <a href="admin_case.asp?id=<%=rs("p_id")%>&action=delpro#delpro">删除</a> <a href="../showcase.asp?p_id=<%=rs("p_id")%>" target="_blank">查看</a></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有案例！</td>
        </tr>
<%End if
End if%>
</table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form100" method="post" action="admin_case.asp">
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
              <a href="admin_case.asp?page=<%=i%>"><%=i%></a> 
              <%End if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
End if%>
              <a href="admin_case.asp?page=<%=page%>" title="下一页">>></a> 
              <%End if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>
      <br>
	  <%if request.QueryString("action")="newpro" then%>
      
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="">
    <INPUT name="upfile" onchange=return(OnUpFile()) type=hidden>
    <tr> 
      <td colspan="2" bgcolor="#E8E8E8"><a name="newpro">新的案例</a></font></td>
    </tr>
    <tr> 
      <td height="25" bgcolor="#FFFFFF" class="chinese">&nbsp;案例名称: 
        <input name="p_name" type="text" class="textarea" id="p_name" size="30" maxlength="30">
        &nbsp;&nbsp; &nbsp; </td>
      <td height="25" bgcolor="#FFFFFF" class="chinese"><input name="p_other2" type="checkbox" id="p_other2" value="1">
        加为推荐案例</td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;案例地址: 
        <input name="p_spec" type="text" class="textarea" id="p_spec" value="" size="30" maxlength="30"> 
        &nbsp;&nbsp; &nbsp; </td>
      <td bgcolor="#FFFFFF" class="chinese">工作范围: 
        <input name="p_fanwei" type="text" class="textarea" id="p_fanwei" size="30" maxlength="30"> 
      </td>
    </tr>
    <tr> 
      <td width="36%" height="25" bgcolor="#FFFFFF" class="chinese">&nbsp;案例类别: 
        <select name="p_class_id" id="p_class_id">
          <option value="" selected>-请选择商品分类-</option>
          <% 
			'###############取出类别分类#################
			dim sql_p,rs_p
			sql_p="select p_id,p_type from case_class"
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
			Do While Not rs_p.EOF 
			  %>
          <option value="<%= rs_p("p_id") %>"><%= rs_p("p_type") %></option>
          <% 
			  rs_p.movenext
			  loop
			  rs_p.close
			  set rs_p=Nothing
			  'conn.close
			  'set conn=Nothing
			  End if
			'####################取出分类结束#####################  
			  %>
        </select> </td>
      <td width="64%" bgcolor="#FFFFFF" class="chinese">工作时间: 
        <input name="p_time" type="text" class="textarea" id="p_time" size="30" maxlength="30"> 
        &nbsp;&nbsp; </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" bgcolor="#FFFFFF" class="chinese"><table width="100%" height="100%" cellpadding="0"  cellspacing="0">
          <tr> 
            <td width="11%">&nbsp;案例图片:</td>
            <td width="89%" align="left" valign="top"><input name="p_pic" type="text" class="textarea" id="p_pic" value="" size="30" maxlength="100"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="25" colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">&nbsp;简单说明： 
        以下编辑器使用和Word相似,但请注意内容不要太多字数在1500以下为佳。</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF"> <textarea name="p_epitome" style="display:none"></textarea> 
        <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=p_epitome&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe> 
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" align="center" bgcolor="#F5F5F5" class="chinese"> 
        <input type="submit" name="Submit2" value="确定新增" class="button"> &nbsp; 
        <input type="reset" name="Reset" value="清空重填" class="button"> </td>
    </tr>
    <input type="hidden" name="action" value="newpro">
    <input type="hidden" name="MM_insert" value="true">
  </form>
</table>
	  <%End if
if request.QueryString("action")="editpro" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>请指定操作的对象！"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>非法的案例ID参数！"
	Call diserror()
	response.End
  End if
End if
sql="select * from case_info where p_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="">
    <INPUT name="upfile" onchange=return(OnUpFile()) type=hidden>
    <input name="upfile2" type="hidden" value="<%= rs("p_pic") %>">
    <tr> 
      <td colspan="2" bgcolor="#E8E8E8"><a name="newpro">修改案例</a></font>&nbsp; 
        (注意：不要修改的地方，请不要改动保持原样！）</td>
    </tr>
    <tr> 
      <td width="36%" bgcolor="#FFFFFF" class="chinese">&nbsp;案例名称: 
        <input name="p_name" type="text" class="textarea" id="p_name" value="<%= rs("p_name") %>" size="30" maxlength="30"> 
        &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;</td>
      <td width="64%" bgcolor="#FFFFFF" class="chinese">
<input name="p_other" type="checkbox" id="p_other3" value="1" <% if rs("p_other")=1 then %>checked <% End if %>>
        加为推荐案例(注意：不选为不推荐)</td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;案例地址: 
        <input name="p_spec" type="text" class="textarea" id="p_spec" value="<%= rs("p_spec") %>" size="30" maxlength="30"> 
        &nbsp;&nbsp; </td>
      <td bgcolor="#FFFFFF" class="chinese">工作范围: 
        <input name="p_fanwei" type="text" class="textarea" id="p_fanwei" value="<%= rs("p_fanwei") %>" size="30" maxlength="30"> 
      </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;案例类别: 
        <select name="p_class_id" id="p_class_id">
          <% 
			'###############取出类别分类#################
			'dim sql_p,rs_p
			sql_p="select p_id,p_type from case_class"
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
			Do While Not rs_p.EOF 
			  %>
          <option value="<%= rs_p("p_id") %>" <% if rs_p("p_id")=rs("p_class_id") then Response.Write("selected") End if %> ><%= rs_p("p_type") %></option>
          <% 
			  rs_p.movenext
			  loop
			  rs_p.close
			  set rs_p=Nothing
			  'conn.close
			  'set conn=Nothing
			  End if
			'####################取出分类结束#####################  
			  %>
        </select> </td>
      <td bgcolor="#FFFFFF" class="chinese">工作时间: 
        <input name="p_time" type="text" class="textarea" id="p_time" value="<%= rs("p_time") %>" size="30" maxlength="30"> 
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" bgcolor="#FFFFFF" class="chinese"><table width="100%" height="100%" cellpadding="0"  cellspacing="0">
          <tr> 
            <td width="9%">&nbsp;案例图片:</td>
            <td width="60%" align="left" valign="top"> <input name="p_pic" type="text" class="textarea" id="p_pic" value="<%= rs("p_pic") %>" size="30" maxlength="100"></td>
            <td width="31%" align="left" valign="middle"><a href="oldpro_show.asp?path=<%= rs("p_pic") %>&pname=<%=rs("p_name")  %>">【查看图片</a>】</td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">&nbsp;简单说明： 
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese"> <textarea name="p_epitome" style="display:none"><%=Server.HTMLEncode(rs("p_epitome"))%></textarea> 
        <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=p_epitome&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe> 
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" align="center" bgcolor="#F5F5F5" class="chinese"> 
        <input type="submit" name="Submit2" value="确定修改" class="button"> &nbsp; 
        <input type="reset" name="Reset" value="清空重填" class="button"> </td>
    </tr>
    <input type="hidden" name="id" value="<%= cint(request.querystring("id")) %>">
    <input type="hidden" name="action" value="editpro">
    <input type="hidden" name="MM_insert" value="true">
  </form>
</table>
	  <%End if
if request.QueryString("action")="delpro" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>请指定操作的对象！"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>非法的案例ID参数！"
	Call diserror()
	response.End
  End if
End if
sql="select * from case_info where p_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="newpro">删除案例</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">&nbsp;案例名称:<%= rs("p_name") %>
              </td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="chinese">&nbsp;案例地址:<%= rs("p_spec") %></td>
          </tr>
          <tr> 
            <td align="left" bgcolor="#FFFFFF" class="chinese">&nbsp;简单说明:</td>
          </tr>
		   <tr> 
            <td align="left" valign="top" bgcolor="#FFFFFF" class="chinese">&nbsp;<%= rs("p_epitome") %></td>
          </tr>
         
          
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="确定删除" class="button">
&nbsp; 【<a href="admin_case.asp">返回</a>】 </td>
          </tr>
		  <input type="hidden" name="id" value="<%= request.querystring("id") %>">
		  <input type="hidden" name="action" value="delpro">
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
End sub%>