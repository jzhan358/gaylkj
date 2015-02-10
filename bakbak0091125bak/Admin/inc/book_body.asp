<%sub book_body()
dim totalnews,Currentpage,totalpages,i
sql="select * from msg order by m_id DESC"
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
          <td align="center" valign="middle" bgcolor="#E8E8E8">留言管理</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">编号</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">标题</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">操作</td>
        </tr>
<%if Not rs.EOF then
dim newsperpage
dim page1
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
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("m_id")%>　</td>
          <td bgcolor="#FFFFFF" class="chinese">[<span class="date"><%=rs("m_date")%></span>]&nbsp;<a href="admin_book.asp?m_id=<%=rs("m_id")%>&action=viewbook"><%=rs("m_title")%></a> 　</td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="admin_book.asp?m_id=<%=rs("m_id")%>&action=viewbook">查看</a> <a href="admin_book.asp?id=<%=rs("m_id")%>&action=delbook">删除</a></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">当前没有留言！</td>
        </tr>
<%End if
End if%>
</table>
    </table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_book.asp">
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
              <a href="admin_book.asp?page=<%=i%>"><%=i%></a> 
              <%End if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
End if%>
              <a href="admin_book.asp?page=<%=page%>" title="下一页">>></a> 
              <%End if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>

	  <%
if request.QueryString("action")="viewbook" then
if request.querystring("m_id")="" then
  errmsg=errmsg+"<br>"+"<li>请指定操作的对象！"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("m_id")) then
    errmsg=errmsg+"<br>"+"<li>非法的ID参数！"
	Call diserror()
	response.End
  End if
End if
sql="select * from msg where m_id="&cint(request.querystring("m_id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if Not(rs.BOF or rs.EOF) then
%>

      <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#e5e5e5">
			<form action="admin_book.asp" method="post" name="form0" target="_self">
			<input name="action" type="hidden" value="delbook">
			<input name="id" type="hidden" value=<%= cint(request.querystring("m_id")) %>>
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
                                    <td height="25" colspan="5"><%= rs("m_city") %>&nbsp;&nbsp;<%= rs("m_addr") %></td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF">
                                    <td align="right" valign="middle">留言内容:</td>
                                    <td colspan="5"><%= formatStr(rs("m_text")) %></td>
                                  </tr>
                                  <tr align="center" valign="middle" bgcolor="#FFFFFF">
                                    <td height="30" colspan="6"><input type="submit" name="Submit" value=" 删除 " class="button">
&nbsp;&nbsp;
<input type="button" name="Submit" value=" 返回 " onClick="javascript:history.back()" class="button"></td>
                                  </tr>

	    </form>
</table>
	  <%
	  End if
	  End if
	  %>
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
