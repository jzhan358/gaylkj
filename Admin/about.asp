<!--#include file="top.asp"-->

<HTML>
<HEAD>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>

<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<script type="text/javascript" charset="utf-8" src="../kindeditor/kindeditor.js"></script>


</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td bgcolor="#E8E8E8">关于我们</font></td>
		  <td bgcolor="#FFFFFF" colspan="2">[<a href="about.asp?language=0">中文栏目</a>][<a href="about.asp?language=1">英文栏目</a>]</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">栏目名称</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
        </tr>
<%
dim totalnews,Currentpage,totalpages,Language
Language=request.QueryString("language")
if Language<>"1" then language="0"
sql="select * from inc_class  where i_language=" & language &" order by i_id ASC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if Not (rs.EOF and rs.bof) then
i=0
Do While Not rs.EOF%>
        <tr> 
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("i_id")%></td>
          <td bgcolor="#FFFFFF" class="chinese"><%=rs("i_type")%></td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><!--<a href="admin_about.asp?id=<%'=rs("i_id")%>&Language=<%'=rs("i_language")%>&action=editabout#editabout">修改</a> 
            <a href="admin_about.asp?id=<%'=rs("i_id")%>&Language=<%'=rs("i_language")%>&action=delabout#delabout">删除</a>--><% Call v1(rs("i_id"),rs("i_type")) %></td>
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

<%if action="newabout" then%>   
<!--增加分类-->      
<br>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="about_save.asp">
	<tr> 
	  <td bgcolor="#E8E8E8"><a name="newabout">新的栏目</a></font></td>
	</tr>
	<tr> 
	  <td bgcolor="#FFFFFF" class="chinese">栏目名称 :
  <input name="i_type" type="text" class="textarea"  size="20">
  版本：
  <input name="i_language" type="radio" value="0" checked>
  中文
  <input type="radio" name="i_language" value="1">
  英文
  
	  </td>
	</tr>
	<tr> 
	  <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
		<input type="submit" name="Submit2" value="确定新增" class="button">
		<input type="reset" name="Reset" value="清空重填" class="button">
	  </td>
	</tr>
	<input type="hidden" name="action" value="newabout">
    <input type="hidden" name="Language" value="<%=Language%>">	
  </form>
</table>
<%End if%>


<%if action="editabout" then
	id=request.QueryString("id")
	if not chkrequest(id) then alert "error","",1
	sql="select * from inc_class where i_id="&id	
	rs.open sql,conn,1,1
%>
<!--编辑分类--> 
<br>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="about_save.asp">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="editabout">修改栏目</a></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">栏目名称: 
              <input type="text" name="i_type" size="20" class="textarea" value="<%=rs("i_type")%>">
              版本：
              
			  <input type="radio" name="i_language" value="0" <%if rs("i_language")="0" then%> checked<%end if%>>
中文
<input type="radio" name="i_language" value="1" <%if rs("i_language")="1" then%> checked<%end if%>>
英文            </td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="hidden" name="id" value="<%=rs("i_id")%>">
              <input type="hidden" name="action" value="editabout">  
              <input type="hidden" name="Language" value="<%=Language%>">	 
              <input type="submit" name="Submit" value="确定修改" class="button"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_about.asp">返回</a>] </td>
          </tr>		         
        </form>
      </table>
<%
rs.close
End if%>

<%
if action="delabout" then
	id=request.QueryString("id")
	if not chkrequest(id) then alert "error","",1
	sql="select * from inc_class where i_id="&id	
	rs.open sql,conn,1,1
%>
<br>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="about_save.asp">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="delabout" id="delabout">删除栏目</a>&nbsp;&nbsp;注意：删除栏目将同时删除栏目内容,建议你使用修改功能！</font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">栏目名称: <%=rs("i_type")%> 
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="hidden" name="id" value="<%=rs("i_id")%>">
              <input type="hidden" name="action" value="delabout">  
              <input type="hidden" name="Language" value="<%=Language%>">  
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除"> 
              [<a href="admin_about.asp">返回</a>] </td>
          </tr>                
        </form>
      </table>
<%
rs.close
End if%>

<%
if action="viewabout" then
	id=request.QueryString("id")
	if not chkrequest(id) then alert "error","",1
	sql="select inc_text from inc_info where inc_class_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
%>
<script type="text/javascript">
		KE.show({
			id : 'inc_text',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true,
			afterCreate : function(id) {
				KE.event.ctrl(document, 13, function() {
					KE.util.setData(id);
					document.forms['form8'].submit();
				});
				KE.event.ctrl(KE.g[id].iframeDoc, 13, function() {
					KE.util.setData(id);
					document.forms['form8'].submit();
				});
			}
		});
		
		function chk(){
		  KE.util.setData('inc_text');		 
		  return true;		 
		}
	</script>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form8" id="form8" method="post" action="about_save.asp">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="viewabout" id="viewabout">查看内容</a></font></td>
          </tr>          
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
            
            <textarea id="inc_text" name="inc_text" style="width:90%;height:500px;visibility:hidden;"><%=replace_t(rs("inc_text"))%></textarea>
             
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="hidden" name="id" value="<%=id%>">		 
              <input type="hidden" name="action" value="edittextabout"> 
              <input type="hidden" name="Language" value="<%=Language%>">  
              <input name="Submit" type="submit" class="button" id="Submit" value="确定修改" onclick="return chk()"> 
              [<a href="admin.asp">返回</a>] </td>
          </tr>                 
        </form>
      </table>
<%
rs.close
End if%>
	  
	  
<%if action="addabout" then
	id=request.QueryString("id")
	i_type=checkstr(request.QueryString(i_type))	
	if not chkrequest(id) then alert "error","",1
%>
      <script type="text/javascript">
		KE.show({
			id : 'inc_text',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true,
			afterCreate : function(id) {
				KE.event.ctrl(document, 13, function() {
					KE.util.setData(id);
					document.forms['form8'].submit();
				});
				KE.event.ctrl(KE.g[id].iframeDoc, 13, function() {
					KE.util.setData(id);
					document.forms['form8'].submit();
				});
			}
		});
		
		function chk(){
		  KE.util.setData('inc_text');		 
		  return true;		 
		}
	</script>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form9" method="post" action="about_save.asp">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="addabout" id="addabout">增加内容</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><%=i_type%>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
            <textarea id="inc_text" name="inc_text" style="width:90%;height:500px;visibility:hidden;"></textarea>
             <!-- <textarea name="inc_text" style="display:none"></textarea>
             <iframe ID="eWebEditor1" src="../eWeb/ewebeditor.htm?id=inc_text&style=coolblue" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>-->
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="hidden" name="id" value="<%=id%>">		  
              <input type="hidden" name="action" value="addabout">
              <input type="hidden" name="Language" value="<%=Language%>"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定增加"> 
              [<a href="admin_about.asp">返回</a>] </td>
          </tr>
                   
        </form>
      </table>
<% 
End if 
set rs=Nothing

sub v1(id,i_type)
v_sql="select inc_class_id from inc_info where inc_class_id="&id
set v_rs=server.createobject("adodb.recordset")
v_rs.open v_sql,conn,1,1
if v_rs.EOF and v_rs.BOF then
%>
<a href="about.asp?id=<%=id%>&Language=<%=Language%>&action=addabout">[增加栏目内容]</a>
<%
else
%>
<a href="about.asp?id=<%=id%>&Language=<%=Language%>&action=viewabout">[修改栏目内容]</a>

<%
End if
v_rs.close
set v_rs=Nothing
End sub
%>