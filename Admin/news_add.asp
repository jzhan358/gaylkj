
<!--#include file="top.asp"-->

<%
id=request.QueryString("id")
cid=request.querystring("cid")
if chkrequest(cid) then cid=clng(cid) else cid=0
action="new"
action_name="新增"
news_title=""
news_keyword=""
news_author=""
news_ahome=""
news_class_id=""
news_content=""
if chkrequest(id) then 
	sql="select * from news where news_id="&id
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	if not rs.eof then
		news_title=rs("news_title")
		news_keyword=rs("news_keyword")
		news_author=rs("news_author")
		news_ahome=rs("news_ahome")
		news_class_id=rs("news_class_id")
		news_content=rs("news_content")
		action="edit"
		action_name="编辑"
	end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>管理中心</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
<script language="javascript" src="../js/validator.js"></script>
<script type="text/javascript" charset="utf-8" src="../kindeditor/kindeditor.js"></script>

<script type="text/javascript">
		KE.show({
			id : 'news_content',
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
		  KE.util.setData('news_content');		 
		  return true;		 
		}
	</script>
</HEAD>
<body>  
<br>  

<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" id="form1" method="post" action="news_save.asp" onSubmit="return Validator.Validate(this,2);">
    <tr> 
      <td bgcolor="#E8E8E8"><a name="newnews"><%=action_name%>信息</a></font></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">标 &nbsp;题:
<input type="text" name="news_title" class="textarea" size="50"  dataType="LimitB" min="2" max="200" msg="标题必须填写（2-200字符）" value="<%=news_title%>">
*
         </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">关键字:
        <input type="text" name="news_keyword" value="<%=news_keyword%>" size="80" class="textarea">
（填写相关热门的关键字使在百度里面更容易找到）
      </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">作 &nbsp;者:
<input name="news_author" type="text" class="textarea" value="<%=news_author%>" size="20">       
      </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">来 &nbsp;源:
<input type="text" name="news_ahome" class="textarea" size="50" value="<%=news_ahome%>">
      </td>
    </tr>
    
    <tr>
      <td height="31" bgcolor="#FFFFFF" class="chinese">选择分类: 
         <select name="news_class_id" id="news_class_id" dataType="Required" msg="没有选择信息分类">         
        <%
      '###############取出类别分类#################
      dim sql_news,rs_news
      sql_news="select news_id,news_type from news_class"
      set rs_news=server.createobject("adodb.recordset")
      rs_news.open sql_news,conn,1,1
      if Not(rs_news.BOF and rs_news.EOF) then 
     	Do While Not rs_news.EOF 
        %>
        <option value="<%= rs_news("news_id") %>" <%if clnG(rs_news("news_id"))=cid and action="new" then response.Write("selected")%><%if clnG(rs_news("news_id"))=news_class_id and action="edit" then response.Write("selected")%>><%=rs_news("news_type") %></option>
        <% 
        rs_news.movenext
        loop       
      end if
	  rs_news.close
      set rs_news=Nothing  
      '####################取出分类结束#####################  
        %>
        </select>
         *(请选择正确的信息分类。如果是英文的信息，请选择英文的分类) </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">
     <textarea id="news_content" name="news_content" style="width:90%;height:500px;visibility:hidden;"><%=news_content%></textarea>
            
       　            </td>
    </tr>
    <tr> 
      <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
      <input type="hidden" name="action" value="<%=action%>">
    <input type="hidden" name="id" value="<%=id%>">
        <input type="submit" name="Submit2" value="提交信息" class="button" onclick="return chk()">
        <input type="reset" name="Reset" value="清空重填" class="button">
      </td>
    </tr>
    
  </form>
</table>
<%
set rs=Nothing
closeconn
%>