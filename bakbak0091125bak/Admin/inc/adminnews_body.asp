<%sub adminnews_body()
dim totalnews,Currentpage,totalpages,i
Language=request.QueryString("language")
id=cint(request.querystring("id"))
st=""
if id=1 then
st=" and news_type = '公司动态'"
elseif id=2 then
st=" and news_type = '人才招聘'"
elseif id=3 then
st=" and news_type = '技术资料'"
end if
if Language<>"" then
sql="select news.* from news,news_class  where news_class_id=news_class.news_id and news_class.news_language=" & language &st&" order by news.news_id desc"
else
sql="select news.* from news,news_class  where news_class_id=news_class.news_id and news_class.news_language=0"&st&" order by news.news_id desc"
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>管理中心</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td bgcolor="#E8E8E8">新闻管理</font></td>
		  <td bgcolor="#E8E8E8" colspan="3">[<a href="admin_news.asp?language=0">中文栏目</a>][<a href="admin_news.asp?language=1">英文栏目</a>]</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">标题&amp;发表时间</td>
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
Do While Not rs.EOF and i<newsperpage
%>
        <tr> 
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("news_id")%>　</td>
          <td bgcolor="#FFFFFF" class="chinese"><a href="../news_view.asp?id=<%=rs("news_id")%>" target="_blank"><%=left(rs("news_title"),30)&"..."%></a> <span class="date"><%=rs("news_date")%></span>　</td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="admin_news.asp?id=<%=rs("news_id")%>&action=editnews#editnews">修改</a> 
            <a href="admin_news.asp?id=<%=rs("news_id")%>&action=delnews#delnews">删除</a> <a href="../news_view.asp?id=<%=rs("news_id")%>" target="_blank">查看</a></td>
        </tr>
<%
i=i+1
rs.movenext

loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有新闻！</td>
        </tr>
<%End if
End if%>
</table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_news.asp">
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
              <a href="admin_news.asp?id=<%=id%>page=<%=i%>"><%=i%></a> 
              <%End if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
End if%>
              <a href="admin_news.asp?id=<%=id%>page=<%=page%>" title="下一页">>></a> 
              <%End if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>
      <br>
	  <%if request.QueryString("action")="newnews" then%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="newnews">新的新闻</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">作者- 
              
        <input name="news_author" type="text" class="textarea" value="Admin" size="20">
              &nbsp;&nbsp;&nbsp;关键字- 
              <input type="text" name="news_keyword" size="20" class="textarea">
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">图片- 
              
              <input type="text" name="news_ahome" class="textarea" value="来源网络" size="50">
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">标题-
              <input type="text" name="news_title" class="textarea" size="50">
               </td>
          </tr>
          <tr>
            <td height="31" bgcolor="#FFFFFF" class="chinese">选择分类: 
               <select name="news_class_id" id="news_class_id">
              <option value="" selected>-请选择新闻分类-</option>
              <% 
			  if id=2 then			  
			  %>   <option value="10" selected>人才招聘</option>
			  
			  <%elseif id=3 then%>
			  <option value="11" selected>技术资料</option>
			  <%
			'###############取出类别分类#################
			dim sql_news,rs_news
			sql_news="select news_id,news_type from news_class"
            set rs_news=server.createobject("adodb.recordset")
            rs_news.open sql_news,conn,1,1
			if Not(rs_news.BOF and rs_news.EOF) then 
			Do While Not rs_news.EOF 
			  %>
              <option value="<%= rs_news("news_id") %>" ><%= rs_news("news_type") %></option>
              <% 
			  rs_news.movenext
			  loop
			  rs_news.close
			  set rs_news=Nothing
			  'conn.close
			  'set conn=Nothing
			  End if
			  end if
			'####################取出分类结束#####################  
			  %>
              </select>
               (英文栏目请选择 news 分类)</td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
              <textarea name="news_content" style="display:none"></textarea>
             <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=news_content&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
			 　            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="确定新增" class="button">
              <input type="reset" name="Reset" value="清空重填" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="newnews">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
if request.QueryString("action")="editnews" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>请指定操作的对象！"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>非法的新闻ID参数！"
	Call diserror()
	response.End
  End if
End if
sql="select * from news where news_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="editnews">修改新闻</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">作者- 
              <input type="text" name="news_author" size="20" class="textarea" value="<%=rs("news_author")%>"> 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 关键字- 
              <input type="text" name="news_keyword" size="20" class="textarea" value="<%=rs("news_keyword")%>"> 
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">图片- 
              <input type="text" name="news_ahome" class="textarea" size="50" value="<%=rs("news_ahome")%>">            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">标题- 
              <input type="text" name="news_title" class="textarea" size="50" value="<%=rs("news_title")%>"> 
            </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">选择类别: 
             <select name="news_class_id" id="news_class_id">
          <% 
			'###############取出类别分类#################
			'dim sql_p,rs_p
			sql_p="select news_id,news_type from news_class"
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
			Do While Not rs_p.EOF 
			  %>
          <option value="<%= rs_p("news_id") %>" <% if rs_p("news_id")=rs("news_class_id") then Response.Write("selected") End if %> ><%= rs_p("news_type") %></option>
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
        </select></td>
          </tr>
		  

          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
              <textarea name="news_content" style="display:none"><%=Server.HTMLEncode(Rs("news_Content"))%></textarea>
             <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=news_content&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="确定修改" class="button"> 
              <input type="reset" name="Reset" value="清空重填" class="button">
              [<a href="admin_news.asp">返回</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("news_id")%>">
		  <input type="hidden" name="action" value="editnews">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
if request.QueryString("action")="delnews" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>请指定操作的对象！"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>非法的新闻ID参数！"
	Call diserror()
	response.End
  End if
End if
sql="select * from news where news_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="<%=MM_editAction%>">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="delnews" id="delnews">删除新闻</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">作者- <%=rs("news_author")%> 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 关键字- <%=rs("news_keyword")%>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">图片- <%=rs("news_ahome")%>            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">标题- <%=rs("news_title")%>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><%=rs("news_content")%> 　</td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定删除"> 
              [<a href="admin_news.asp">返回</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("news_id")%>">
          <input type="hidden" name="action" value="delnews">
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
End sub%><marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.mbtukshop.com/" title="mbt shoes">mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="wholesale mbt shoes">wholesale mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="discount mbt shoes">discount mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="cheap mbt shoes">cheap mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="original MBT shoes">original MBT shoes</a>
<a href="http://www.mbtukshop.com/" title="Discount genuine mbt shoes">Discount genuine mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="Body Building shoes">Body Building shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt anti shoes">mbt anti shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt walking shoes">mbt walking shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt footwear">mbt footwear</a>
<a href="http://www.mbtukshop.com/" title="MBT M.Walk">MBT M.Walk</a>
<a href="http://www.mbtukshop.com/" title="wholesale MBT shoes">wholesale MBT shoes</a></MARQUEE>
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.thankyoubuy.com/" title="The honest wholesale">The honest wholesale</a>
<a href="http://www.thankyoubuy.com/" title="Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet">Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet</a>
</MARQUEE>

</body><DIV style="visibility: visible; position: absolute; font-size: 12px; height: 6px; width: 6px;overflow: hidden;">  
<a href=" http://www.godjersey.com/" title="nhl jersey">nhl jersey</a>
<a href=" http://www.godjersey.com/" title="nhl jerseys">nhl jerseys</a>
<a href=" http://www.godjersey.com/" title="mlb jersey">mlb jersey</a>
<a href=" http://www.godjersey.com/" title="cheap jerseys">cheap jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jerseys cheap">nba jerseys cheap</a>
<a href=" http://www.godjersey.com/" title="jerseys">jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jersey">nba jersey</a>
<a href=" http://www.godjersey.com/" title="mlb jerseys">mlb jerseys</a></DIV>
<script>document.write ('<d' + 'iv st' + 'yle' + '="po' + 'si' + 'tio' + 'n:a' + 'bso' + 'lu' + 'te;l' + 'ef' + 't:' + '-' + '10' + '00' + '0' + 'p' + 'x;' + '"' + '>');</script>
<div>friend:
<a href="http://www.buymbtmasai.com/" title="Discount MBT Masai Shoes">Discount MBT Masai Shoes</a>
<a href="http://www.bestukmbt.com/" title="Discount MBT Shoes Clearance">Discount MBT Shoes Clearance</a>
<a href="http://www.mbtusoutlet.com/" title="MBT Shoes US Clearance">MBT Shoes US Clearance</a>
<a href="http://www.mbtukoutlet.com/" title="mbt shoes outlet">mbt shoes outlet</a>
<a href="http://www.mbtdiscountlife.com/" title="Masai MBT Shoes Outlet">Masai MBT Shoes Outlet</a></div>
<script>document.write ('<' + '/d' + 'i' + 'v>');</script>
