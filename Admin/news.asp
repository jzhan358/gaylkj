
<!--#include file="top.asp"-->
<!--#include file="../inc/page.asp"-->
<%
dim totalnews,Currentpage,totalpages,i
dim language,cid,keyword
language=request.QueryString("language")
cid=request.querystring("cid")
st=""
if chkrequest(cid) then 
st=" and news_class_id="&cid 
cid=clng(cid)
else
cid=0
end if
if language="0" or language="1" then st=" and news_language="&language
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
</HEAD>
<body background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td bgcolor="#E8E8E8">信息管理</font></td>
		  <td bgcolor="#E8E8E8" colspan="3">&nbsp;<!--[<a href="news.asp?language=0&cid=<%=cid%>">中文信息</a>][<a href="news.asp?language=1&cid=<%=cid%>">英文信息</a>]--></td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">标题&amp;发表时间</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
        </tr>
<%


sql="select news_id,news_title,news_date from news where news_id>0"&st

dim sql_id,sql_Field,sql_table,sql_order,sqlst
dim pagesize,pagecount,recordcount,page
'条件
sqlst="news_id>0"&st
'统计字段   
sql_id="news_id"  
'显示字段   
sql_Field="news_id,news_title,news_date"  
'查询表名   
sql_table="news"  
'排序字段   
sql_order="news_id"  
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
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("news_id")%>　</td>
          <td bgcolor="#FFFFFF" class="chinese"><a href="../news_view.asp?id=<%=rs("news_id")%>" target="_blank"><%=rs("news_title")%></a> [<span class="date"><%=rs("news_date")%></span>]　</td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="news_add.asp?id=<%=rs("news_id")%>&cid=<%=cid%>&page=<%=page%>&keyword=<%=keyword%>&language=<%=language%>">修改</a> 
            <a href="news.asp?action=del&id=<%=rs("news_id")%>&page=<%=page%>&keyword=<%=keyword%>&cid=<%=cid%>&language=<%=language%>">删除</a> <a href="../news_view.asp?id=<%=rs("news_id")%>" target="_blank">查看</a></td>
        </tr>
<%
i=i+1
rs.movenext

loop
else
%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有信息！</td>
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

<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="news_save.asp" onSubmit="return Validator.Validate(this,2);">
    <tr> 
      <td bgcolor="#E8E8E8"><a name="newnews">新的信息</a></font></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">标 &nbsp;题:
<input type="text" name="news_title" class="textarea" size="50"  dataType="LimitB" min="2" max="200" msg="标题必须填写（2-200字符）">
*
         </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">关键字:
        <input type="text" name="news_keyword" size="80" class="textarea">
（填写相关热门的关键字使在百度里面更容易找到）
      </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">作 &nbsp;者:
<input name="news_author" type="text" class="textarea" value="Admin" size="20">       
      </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">来 &nbsp;源:
<input type="text" name="news_ahome" class="textarea" value="Internet" size="50">
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
        <option value="<%= rs_news("news_id") %>" <%if clnG(rs_news("news_id"))=cid then response.Write("selected")%>><%=rs_news("news_type") %></option>
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
      
             <textarea id="news_content" name="news_content" cols="100" rows="8" style="width:98%;height:300px;visibility:hidden;"></textarea>
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
sql="select * from news where news_id="&id
rs.open sql,conn,1,1
%>
      
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="news_save.asp?page=<%=page%>&keyword=<%=keyword%>&cid=<%=cid%>&language=<%=language%>">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="editnews">修改信息</a></font></td>
          </tr>
           <tr> 
            <td bgcolor="#FFFFFF" class="chinese">标题: 
              <input type="text" name="news_title" class="textarea" size="50" value="<%=rs("news_title")%>" dataType="LimitB" min="2" max="200" msg="标题必须填写（2-200字符）"> 
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">关键字: 
              <input type="text" name="news_keyword" size="80" class="textarea" value="<%=rs("news_keyword")%>"> 
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">作者: 
              <input type="text" name="news_author" size="20" class="textarea" value="<%=rs("news_author")%>">              
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">来源: 
              <input type="text" name="news_ahome" class="textarea" size="50" value="<%=rs("news_ahome")%>">            </td>
          </tr>
         
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">选择类别: 
             <select name="news_class_id" id="news_class_id" dataType="Required" msg="没有选择信息分类">
          <% 
			'###############取出类别分类#################
			'dim sql_p,rs_p
			sql_p="select news_id,news_type from news_class"
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
				Do While Not rs_p.EOF 
				  %>
			  <option value="<%= rs_p("news_id") %>" <% if clng(rs_p("news_id"))=clng(rs("news_class_id")) then Response.Write("selected") End if %> ><%= rs_p("news_type") %></option>
			  <% 
				rs_p.movenext
				loop				  	  			
			  End if
			  rs_p.close
			  set rs_p=Nothing
			'####################取出分类结束#####################  
			  %>
        </select>
             (请选择正确的信息分类。如果是英文的信息，请选择英文的分类)</td>
          </tr>
		  

          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
           
              <textarea id="news_content" name="news_content" cols="100" rows="8" style="width:500px;height:300px;visibility:hidden;"><%=rs("news_content")%></textarea>
           
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="确定修改" class="button"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="news.asp">返回</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("news_id")%>">
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
sql="select news_id,news_title from news where news_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="news_save.asp?page=<%=page%>&keyword=<%=keyword%>&cid=<%=cid%>&language=<%=language%>">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="delnews" id="delnews">删除信息</a></font></td>
          </tr>                  
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">标题- <%=rs("news_title")%>
            </td>
          </tr>          
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除"> 
              [<a href="news.asp">返回</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("news_id")%>">
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