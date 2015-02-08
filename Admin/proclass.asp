
<!--#include file="top.asp"-->
<HTML>
<HEAD>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>

<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td bgcolor="#E8E8E8">产品分类</font></td>
		  <td bgcolor="#FFFFFF" colspan="2"></td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">栏目中文名称/英文名称</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
        </tr>
<%
dim totalnews,Currentpage,totalpages,Language

sql="select p_id,p_type,p_type_e from p_class order by p_id desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if Not (rs.EOF and rs.bof) then
i=0
Do While Not rs.EOF%>
        <tr> 
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("p_id")%></td>
          <td bgcolor="#FFFFFF" class="chinese"><%=rs("p_type")%> / <%=rs("p_type_e")%></td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="proclass.asp?id=<%=rs("p_id")%>&action=edit">修改</a> 
            <a href="proclass.asp?id=<%=rs("p_id")%>&action=del">删除</a></td>
        </tr>
<%
i=i+1
rs.movenext
loop
else
%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有栏目！</td>
        </tr>
<%End if%>
</table>   
<%
rs.close
action=request.QueryString("action")
%>

<%if action="new" then%>   
<!--增加分类-->      
<br>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="proclass_save.asp">
	<tr> 
	  <td bgcolor="#E8E8E8"><a name="newabout">新的栏目</a></font></td>
	</tr>
	<tr> 
	  <td bgcolor="#FFFFFF" class="chinese">栏目名称 :
  <input name="p_type" type="text" class="textarea"  size="20">&nbsp;
  英文名称 :
  <input name="p_type_e" type="text" class="textarea"  size="20">
  
	  </td>
	</tr>
	<tr> 
	  <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
		<input type="submit" name="Submit2" value="确定新增" class="button">
		<input type="reset" name="Reset" value="清空重填" class="button">
	  </td>
	</tr>
	<input type="hidden" name="action" value="new">   
  </form>
</table>
<%End if%>


<%if action="edit" then
	id=request.QueryString("id")
	if not chkrequest(id) then alert "error","",1
	sql="select * from p_class where p_id="&id	
	rs.open sql,conn,1,1
%>
<!--编辑分类--> 
<br>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="proclass_save.asp">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="editabout">修改栏目</a></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">栏目名称: 
              <input type="text" name="p_type" size="20" class="textarea" value="<%=rs("p_type")%>">
              &nbsp;
  英文名称 :
  <input name="p_type_e" type="text" class="textarea" value="<%=rs("p_type_e")%>"  size="20">  </td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="hidden" name="id" value="<%=rs("p_id")%>">
              <input type="hidden" name="action" value="edit">               
              <input type="submit" name="Submit" value="确定修改" class="button"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_proclass.asp">返回</a>] </td>
          </tr>		         
        </form>
      </table>
<%
rs.close
End if%>

<%
if action="del" then
	id=request.QueryString("id")
	if not chkrequest(id) then alert "error","",1
	sql="select * from p_class where p_id="&id	
	rs.open sql,conn,1,1
%>
<br>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="proclass_save.asp">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="delabout" id="delabout">删除栏目</a>&nbsp;&nbsp;注意：删除栏目将同时删除栏目内容,建议你使用修改功能！</font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">栏目中文名称: <%=rs("p_type")%> - 栏目英文名称: <%=rs("p_type_e")%>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="hidden" name="id" value="<%=rs("p_id")%>">
              <input type="hidden" name="action" value="del"> 
              
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除"> 
              [<a href="admin_proclass.asp">返回</a>] </td>
          </tr>                
        </form>
      </table>
<%
rs.close
End if%>