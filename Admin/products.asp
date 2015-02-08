
<!--#include file="top.asp"-->

<!--#include file="../inc/page.asp"-->
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8"><LINK 
href="inc/djcss.css" type=text/css rel=StyleSheet>
<script language="javascript" src="../js/validator.js"></script>
<script type="text/javascript" charset="utf-8" src="../kindeditor/kindeditor.js"></script>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<SCRIPT language=javascript>
  var upfile_obj;
  function OnUpFile()
  {
   upfile_obj.src=upfile_obj.lowsrc;
   upfile_obj.src=document.forms["form1"].upfile.value;
  }
</SCRIPT>
<script language = "javascript">
var i,j;
j=0;
goaler = new Array();
<%set rs_p=conn.execute("select * from p_class_small order by p_small_id")
if rs_p.eof then%>
goaler[0] = new Array("无分类","","");
<%else
i=0
do while not rs_p.eof%>
goaler[<%=i%>] = new Array("<%=rs_p("p_small_type")%>","<%=rs_p("p_type")%>");
<%rs_p.movenext
i=i+1
loop
end if
rs_p.close
%>
j=<%=i%>;

function changelocation(location_id)//传递一级分类的值,从而改变二级分类
{
document.form1.p_small_type.length = 0; 

var p_type=location_id;
var i;
for (i=0;i < j; i++)
{
if (goaler[i][1] ==p_type) 
document.form1.p_small_type.options[document.form1.p_small_type.length] = new Option(goaler[i][0],goaler[i][0]);

}

} 
</script>

<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <tr> 
    <td bgcolor="#E8E8E8" colspan="3">产品管理(图片上传时请注意图片的大小和格式，只可以上传大小在200K以下的“.jpg”或“.gif”结尾的图片。)</td>
  </tr>
  <tr bgcolor="#E8E8E8" align="center"> 
    <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
    <td bgcolor="#F5F5F5" class="chinese">产品名称</td>
    <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
  </tr>
  <%
dim sql_id,sql_Field,sql_table,sql_order,sqlst
dim pagesize,pagecount,recordcount,page
'条件
sqlst="p_id>0"&st
'统计字段   
sql_id="p_id"  
'显示字段   
sql_Field="p_id,p_name,p_name_e,p_date"  
'查询表名   
sql_table="p_info"  
'排序字段   
sql_order="p_id"  
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
    <td bgcolor="#FFFFFF" class="chinese" align="center"><%= rs("p_id") %>　</td>
    <td bgcolor="#FFFFFF" class="chinese"><%= rs("p_name") %>/<%= rs("p_name_e") %></td>
    <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="products_edit.asp?id=<%=rs("p_id")%>&action=edit">修改</a> <a href="products_edit.asp?id=<%=rs("p_id")%>&action=del">删除</a> </td>
  </tr>
  <%i=i+1
rs.movenext
loop
else
%>
  <tr align="center"> 
    <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有产品！</td>
  </tr>
  <%
End if%>
</table>
<table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">	   
  <tr>
    <td class="chinese" align="right"><%=showpage(pagecount,pagesize,page,recordcount,15)%></td>
  </tr>	
</table>
      
<%
rs.close
action=request.QueryString("action")
if action="new" then%>
<br>
<script type="text/javascript">
		KE.show({
			id : 'p_jianjie',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true,
			afterCreate : function(id) {
				KE.event.ctrl(document, 13, function() {
					KE.util.setData(id);
					document.forms['form1'].submit();
				});
				KE.event.ctrl(KE.g[id].iframeDoc, 13, function() {
					KE.util.setData(id);
					document.forms['form1'].submit();
				});
			}
		});
		
		KE.show({
			id : 'p_jianjie_e',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true,
			afterCreate : function(id) {
				KE.event.ctrl(document, 13, function() {
					KE.util.setData(id);
					document.forms['form1'].submit();
				});
				KE.event.ctrl(KE.g[id].iframeDoc, 13, function() {
					KE.util.setData(id);
					document.forms['form1'].submit();
				});
			}
		});
		
		function chk(){
		  KE.util.setData('p_jianjie');	
		  KE.util.setData('p_jianjie_e');	
		  return true;		 
		}
	</script>     
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="products_save.asp"  onSubmit="return Validator.Validate(this,2);">
    <INPUT name="upfile" onchange=return(OnUpFile()) type=hidden>
    <tr> 
      <td colspan="2" bgcolor="#E8E8E8"><a name="newpro">新的产品</a></font></td>
    </tr>
    <tr> 
      <td width="36%" height="25" bgcolor="#FFFFFF" class="chinese">&nbsp;产品名称: 
        <input name="p_name" type="text" class="textarea" id="p_name" size="30" maxlength="50" dataType="LimitB" min="2" max="50" msg="产品中文名称不能为空(2-50字)"> 
        &nbsp;<span class="style1">*</span> &nbsp; </td>
      <td width="64%" bgcolor="#FFFFFF" class="chinese">英文名称: 
        <input name="p_name_e" type="text" class="textarea" id="p_name_e" size="30" maxlength="100" dataType="LimitB" min="2" max="150" msg="产品英文名称不能为空(2-150字)">
        <span class="style1">*</span></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;产品型号: 
        <input name="p_spec" type="text" class="textarea" id="p_spec" dataType="LimitB" min="2" max="30" msg="产品中文型号名称不能为空(2-30字符)" value="" size="30" maxlength="30"> 
        &nbsp;<span class="style1">*</span> </td>
      <td bgcolor="#FFFFFF" class="chinese">英文型号: 
        <input name="p_spec_e" type="text" class="textarea" id="p_spec_e" dataType="LimitB" min="2" max="30" msg="产品英文型号名称不能为空(2-30字符)" value="" size="30" maxlength="30">
        <span class="style1">*</span></td>
    </tr>
    <tr> 
      <td height="12" bgcolor="#FFFFFF" class="chinese">&nbsp;产品类别: 
       <!--<select name="p_type" size="1" id="p_type" onChange="changelocation(document.form1.p_type.options[document.form1.p_type.selectedIndex].value)">-->
		<select name="p_type_id" size="1" id="p_type_id" dataType="Require" msg="请选择产品分类">
          <%set rs=conn.execute("select p_id,p_type from p_class")
          if rs.eof then%>
          <option selected value="">无一级分类</option>
          <%else%>
          <option selected value="">一级分类</option>
          <%do while not rs.eof
		  %>
          <option value="<%=rs("p_id")%>"><%=rs("p_type")%></option>
          <%rs.movenext
             loop
          end if
		  rs.close
		  %>
        </select> <!--<select name="p_small_type" >
          <option selected value="">二级分类</option>
        </select> -->
        <span class="style1">*</span></td>
      <td height="12" bgcolor="#FFFFFF" class="chinese"><!--<input name="p_other2" type="checkbox" id="p_other2" value="1">
        加为会员可查看(注意：不选为所有人都可以查看)-->
        排序数值:
        <input name="p_order" type="text" class="textarea" id="p_order" value="0" size="15" maxlength="15">
可以不填,数值越大,排得越前(首页滚动前20条产品信息)</td>
    </tr>
  <tr> 
      <td height="15" colspan="2" bgcolor="#FFFFFF" align="left">产品图片:<input name="p_pic" type="text" class="textarea" id="p_pic"  value="" size="30" maxlength="100" datatype="LimitB" min="10" max="150" msg="产品图片不能为空">
        上传图片：<iframe name=uploadformu frameborder=0 width=400 height=22 scrolling=no src=admin_upload.asp?id=p_pic></iframe><font color="#FF0000">(尺寸:500*500左右)</font></td>
    </tr> 
    <tr> 
      <td height="25" colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">&nbsp;中文简介： 
        以下编辑器使用和Word相似,但请注意内容不要太多字数在1500以下为佳。</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF">
      <textarea id="p_jianjie" name="p_jianjie" style="width:90%;height:500px;visibility:hidden;"></textarea>
     
         <!--<textarea name="p_jianjie" style="display:none"></textarea><iframe ID="eWebEditor1" src="../eWeb/ewebeditor.htm?id=p_jianjie&style=coolblue" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe> -->      
      
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF"> 英文简介：</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF">
      <textarea id="p_jianjie_e" name="p_jianjie_e" style="width:90%;height:500px;visibility:hidden;"></textarea>
        
      <!--<textarea name="p_jianjie_e" style="display:none"></textarea><iframe ID="eWebEditor1" src="../eWeb/ewebeditor.htm?id=p_jianjie_e&style=coolblue" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe> -->
        </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" align="center" bgcolor="#F5F5F5" class="chinese"> 
        <input type="submit" name="Submit2" value="确定新增" class="button" onClick="return chk()"> &nbsp; 
        <input type="reset" name="Reset" value="清空重填" class="button"> </td>
    </tr>
    <input type="hidden" name="action" value="new">  
  </form>
</table>
<%End if%>

<%
if action="edit" then
id=request.QueryString("id")
if not chkrequest(id) then alert "error","",1
sql="select * from p_info where p_id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<script type="text/javascript">
		KE.show({
			id : 'p_jianjie',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true,
			afterCreate : function(id) {
				KE.event.ctrl(document, 13, function() {
					KE.util.setData(id);
					document.forms['form1'].submit();
				});
				KE.event.ctrl(KE.g[id].iframeDoc, 13, function() {
					KE.util.setData(id);
					document.forms['form1'].submit();
				});
			}
		});
		
		KE.show({
			id : 'p_jianjie_e',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true,
			afterCreate : function(id) {
				KE.event.ctrl(document, 13, function() {
					KE.util.setData(id);
					document.forms['form1'].submit();
				});
				KE.event.ctrl(KE.g[id].iframeDoc, 13, function() {
					KE.util.setData(id);
					document.forms['form1'].submit();
				});
			}
		});
		
		function chk(){
		  KE.util.setData('p_jianjie');	
		  KE.util.setData('p_jianjie_e');	
		  return true;		 
		}
	</script>  
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="products_save.asp" onSubmit="return Validator.Validate(this,2);">
    <INPUT name="upfile" onchange=return(OnUpFile()) type=hidden>
    <input name="upfile2" type="hidden" value="<%=rs("p_pic")%>">
    <tr> 
      <td colspan="2" bgcolor="#E8E8E8"><a name="newpro">修改产品</a></font>&nbsp; 
        (注意：不修改的地方，请不要改动保持原样！）</td>
    </tr>
    <tr> 
      <td width="38%" bgcolor="#FFFFFF" class="chinese">&nbsp;产品名称: 
        <input name="p_name" type="text" class="textarea" id="p_name" value="<%= rs("p_name") %>" size="30" maxlength="50" dataType="LimitB" min="2" max="50" msg="产品中文名称不能为空(2-50字)"> 
        &nbsp;<span class="style1">*</span>&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; </td>
      <td width="62%" bgcolor="#FFFFFF" class="chinese">英文名称: 
        <input name="p_name_e" type="text" class="textarea" id="p_name_e" value="<%= rs("p_name_e") %>" size="30" maxlength="100" dataType="LimitB" min="2" max="150" msg="产品英文名称不能为空(2-150字)">
        <span class="style1">*</span></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;产品型号: 
        <input name="p_spec" type="text" class="textarea" id="p_spec" value="<%= rs("p_spec") %>" size="30" maxlength="30" dataType="LimitB" min="2" max="30" msg="产品中文型号不能为空(2-30字)"> 
        &nbsp;<span class="style1">*</span> </td>
      <td bgcolor="#FFFFFF" class="chinese">英文型号: 
        <input name="p_spec_e" type="text" class="textarea" id="p_spec_e" value="<%= rs("p_spec_e") %>" size="30" maxlength="30" dataType="LimitB" min="2" max="30" msg="产品英文型号不能为空(2-30字)">
        <span class="style1">*</span></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;产品类别: 
        <select name="p_type_id" id="p_type_id" dataType="Require" msg="请选择产品的分类">
          <% 
			'###############取出类别分类#################
			'dim sql_p,rs_p
			sql_p="select p_id,p_type from p_class"
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
			Do While Not rs_p.EOF
			  %>
          <option value="<%= rs_p("p_id") %>" <% if rs_p("p_id")=rs("p_type_id") then Response.Write("selected") End if %> ><%= rs_p("p_type") %></option>
          <%
			  rs_p.movenext
			  loop			 
			  End if
			  rs_p.close
			  set rs_p=Nothing
			'####################取出分类结束#####################  
			  %>
        </select> <!--<select name="p_small_type">
          <option selected value="<%'=rs("p_small_type")%>"><%'=rs("p_small_type")%></option>
        </select>--> </td>
     <td height="30" bgcolor="#FFFFFF" class="chinese">排序数值: 
        <input name="p_order" type="text" class="textarea" id="p_order" value="<%= rs("p_order") %>" size="15" maxlength="15">
        可以不填,数值越大,排得越前(首页滚动前20条产品信息)</td>
    </tr>
	 <tr> 
      <td height="15" colspan="2" bgcolor="#FFFFFF" align="left">产品图片:<input name="p_pic" type="text" class="textarea" id="p_pic"  value="<%= rs("p_pic") %>" size="30" maxlength="150" dataType="LimitB" min="10" max="150" msg="产品的图片不能为空">
        上传图片：<iframe name=uploadformu frameborder=0 width=400 height=22 scrolling=no src=admin_upload.asp?id=p_pic></iframe><font color="#FF0000">(尺寸：500*500左右)</font></td>
    </tr> <!-- <tr> 
      <td height="15" colspan="2" bgcolor="#FFFFFF" align="left">产品大图:<input name="p_pic2" type="text" class="textarea" id="p_pic2"  value="<%= rs("p_pic2") %>" size="30" maxlength="100">
        上传图片：<iframe name=uploadformu frameborder=0 width=400 height=22 scrolling=no src=admin_upload.asp?id=p_pic2></iframe><font color="#FF0000">(尺寸:500*500左右)</font> </td>
    </tr>  -->
    
    <tr>
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">&nbsp;中文简介：</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">
      <textarea id="p_jianjie" name="p_jianjie" style="width:90%;height:500px;visibility:hidden;"><%=replace_t(rs("p_jianjie"))%></textarea>
        
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese"> 英文简介：</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">
      <textarea id="p_jianjie_e" name="p_jianjie_e" style="width:90%;height:500px;visibility:hidden;"><%=replace_t(rs("p_jianjie_e"))%></textarea>
      
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" align="center" bgcolor="#F5F5F5" class="chinese"> 
        <input type="submit" name="Submit2" value="确定修改" class="button" onClick="return chk()"> &nbsp; 
        <input type="reset" name="Reset" value="清空重填" class="button"> </td>
    </tr>
    <input type="hidden" name="id" value="<%=id%>">
    <input type="hidden" name="action" value="edit">    
  </form>
</table>
<%
rs.close
End if%>


<%
if action="del" then
id=request.QueryString("id")
if not chkrequest(id) then alert "error","",1
sql="select p_name,p_name_e from p_info where p_id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
<form name="form1" method="post" action="products_save.asp">
  <tr> 
    <td bgcolor="#E8E8E8"><a name="newpro">删除产品</a></font></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" class="chinese">标题:<%=rs("p_name")%> - <%=rs("p_name_e")%>
    </td>
  </tr>    
   <tr> 
    <td align="center" valign="top" bgcolor="#FFFFFF" class="chinese">&nbsp;<span class="style1">你 确 定 要 删 除 ？</span></td>
  </tr>
 
  
  <tr> 
    <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
      <input type="submit" name="Submit2" value="确定删除" class="button">
&nbsp; 【<a href="products.asp">返回</a>】 </td>
  </tr>
  <input type="hidden" name="id" value="<%=id%>">
  <input type="hidden" name="action" value="del">		  
</form>
</table>
<%
rs.close
End if%>
      
<%
set rs=Nothing
closeconn
%>