﻿
<%sub p_class_body()
dim totalp,Currentpage,totalpages,i
sql="select * from p_class order by p_id ASC"
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
.style1 {color: #FF0000}
-->
</style>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<br>
<table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td width="12%" bgcolor="#E8E8E8"　width="10%">栏目分类管理</font></td>
          <td width="18%" bgcolor="#FFFFFF"　width="10%"><div align="center"><a href="admin_p_class.asp?action=new" class="style1"> 添加一级分类</a></div></td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td colspan="2" bgcolor="#F5F5F5" class="chinese">栏　目　名　称</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">操　作　选　项</td>
        </tr>
<%if Not rs.EOF then
dim pperpage
pperpage=5
rs.movefirst
rs.pagesize=pperpage
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
   totalp=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*pperpage<totalp then
	       rs.move(currentpage-1)*pperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   End if
   End if
   if (totalp mod pperpage)=0 then
      totalpages=totalp\pperpage
   else
      totalpages=totalp\pperpage+1
   End if
i=0
Do While Not rs.EOF and i<pperpage%>
        <tr> 
          
    <td colspan="2" bgcolor="#F5F5F0" class="chinese">　　　<img src="images/tree_1.gif" width="15" height="15">　&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>免费版本无此功能！</strong></font></td>
          <td align="center" bgcolor="#F5F5F0" class="chinese">
		  <a href="admin_p_class.asp?p_type=<%=rs("p_type")%>&action=p#p">添加二级分类</a> |
		  <a href="admin_p_class.asp?id=<%=rs("p_id")%>&action=edit#edit">修改</a> |
          <a href="admin_p_class.asp?id=<%=rs("p_id")%>&action=del#del">删除</a></td>
        </tr>
      <%
	  set rsSmallClass=server.CreateObject("adodb.recordset")
	  rsSmallClass.open "Select * From p_class_small Where p_type='" & rs("p_type") & "'",conn,1,1
	  if not(rsSmallClass.bof and rsSmallClass.eof) then
	  Do While Not rsSmallClass.eof
	  %>
		<tr>
          
    <td colspan="2" bgcolor="#FFFFFF" class="chinese">　　　　<img src="images/tree_2.gif" width="15" height="15">　</td>
          <td align="center" bgcolor="#FFFFFF" class="chinese">
		  　　　　　　　<a href="admin_p_class.asp?id=<%=rsSmallClass("p_small_id")%>&action=edits#edits">修改</a> |
          <a href="admin_p_class.asp?id=<%=rsSmallClass("p_small_id")%>&action=dels#dels">删除</a></td>
        </tr>
        <%
		rsSmallClass.movenext
		loop
	    end if
	    rsSmallClass.close
	    set rsSmallClass=nothing	
	    i=i+1
		rs.movenext
	    loop
		else
if rs.EOF and rs.BOF then
%>
        <tr align="center"> 
          
    <td bgcolor="#FFFFFF" colspan="3" class="chinese"><font color="#FF0000"><strong>免费版本无此功能！</strong></font></td>
        </tr>
<%End if
End if%>
</table>
      <table align="center" width="80%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_p_class.asp">
        <tr>
          
      <td class="chinese" align="right"> /页,条记录/篇每页. <a href="admin_p_class.asp?page=<%=page%>" title="下一页">>></a> 
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
</table>
      <br>
	  <%if request.QueryString("action")="new" then%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="add" id="add">添加栏目大类</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">分类名称 :
        <input name="p_type" type="text" class="textarea"  size="20">
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="new">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
if request.QueryString("action")="edit" then
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
sql="select * from p_class where p_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="edit" id="edit">修改栏目大类</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">分类名称: 
              <input type="text" name="p_type" size="20" class="textarea" value="<%=rs("p_type")%>"> 
            </td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="确定修改" class="button"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_p_class.asp">返回</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("p_id")%>">
		  <input type="hidden" name="action" value="edit">
          <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
if request.QueryString("action")="del" then
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
sql="select * from p_class where p_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="del" id="del">删除栏目大类</a>&nbsp;&nbsp;注意：删除分类将同时删除分类内容,建议你使用修改功能！</font></td>
          </tr>
          <tr> 
            
      <td bgcolor="#FFFFFF" class="chinese">分类名称: </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除"> 
              [<a href="admin_p_class.asp">返回</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("p_id")%>">
          <input type="hidden" name="action" value="del">
          <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if%>
      	
       <%'这里用来对二级分类进行添加，修改，删除管理操作！
	   if request.QueryString("action")="p" then%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="adds" id="adds">添加二级分类</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">所属大类：
			<select name="p_type" id="p_type">
				<% 
			'###############取出类别分类#################
			dim sql_p,rs_p
			p_type=request.QueryString("p_type")
			sql_p="select p_type from p_class"
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
			Do While Not rs_p.EOF 
			  %>
                <option value="<%= rs_p("p_type") %>" <% if rs_p("p_type")=p_type then Response.Write("selected") End if %>><%= rs_p("p_type") %></option>
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
              </select>
			</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">分类名称 :
            <input name="p_small_type" type="text" class="textarea" id="p_small_type"  size="20"></td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="p">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
if request.QueryString("action")="edits" then
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
sql="select * from p_class_small where p_small_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="edits" id="edits">修改二级分类</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">所属大类:
              
              <select name="p_type" id="p_type">
                <% 
			'###############取出类别分类#################
			dim sql_1,rs_1
			sql_1="select p_type from p_class"
            set rs_1=server.createobject("adodb.recordset")
            rs_1.open sql_1,conn,1,1
			if Not(rs_1.BOF and rs_1.EOF) then 
			Do While Not rs_1.EOF 
			  %>
                <option value="<%= rs_1("p_type") %>" <% if rs_1("p_type")=rs("p_type") then Response.Write("selected") End if %>><%= rs_1("p_type") %></option>
                <% 
			  rs_1.movenext
			  loop
			  rs_1.close
			  set rs_1=Nothing
			  'conn.close
			  'set conn=Nothing
			  End if
			'####################取出分类结束#####################  
			  %>
              </select></td>
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">分类名称:
            <input name="p_small_type" type="text" class="textarea" id="p_small_type" value="<%=rs("p_small_type")%>" size="20"></td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="确定修改" class="button"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_p_class.asp">返回</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("p_small_id")%>">
		  <input type="hidden" name="action" value="edits">
          <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
if request.QueryString("action")="dels" then
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
sql="select * from p_class_small where p_small_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="dels" id="dels">删除二级分类</a>&nbsp;&nbsp;注意：删除分类将同时删除分类内容,建议你使用修改功能！</font></td>
          </tr>
          <tr> 
            
      <td bgcolor="#FFFFFF" class="chinese">所属大类:　分类名称: </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除"> 
              [<a href="admin_p_class.asp">返回</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("p_small_id")%>">
          <input type="hidden" name="action" value="dels">
          <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
%>
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