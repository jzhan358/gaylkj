<%sub p_small_body()
dim totalnews,Currentpage,totalpages,i
sql="select * from p_class_small order by p_small_id ASC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>

<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type="text/css">
A{TEXT-DECORATION: none}
A:link {COLOR: #666666; FONT-FAMILY: ����; TEXT-DECORATION: none}
A:visited {COLOR: #666666; FONT-FAMILY: ����; TEXT-DECORATION: none}
A:active {FONT-FAMILY: ����; TEXT-DECORATION: none}
A:hover {BORDER-BOTTOM: 1px dotted; BORDER-LEFT-WIDTH: 1px; BORDER-RIGHT-WIDTH: 1px; BORDER-TOP-WIDTH: 1px; COLOR: #ff6600; TEXT-DECORATION: none}
BODY {
FONT-SIZE: 12px;
COLOR: #666666;
FONT-FAMILY:  ����;
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
TABLE {BORDER-COLLAPSE: collapse; FONT-FAMILY: ����; FONT-SIZE: 9pt}
.button{height:18px;width:62px;background:#f6f6f9 url(images/ButtonBg.gif); border:solid 1px #5589AA;color: #000000 ;FONT-SIZE: 9pt}
.lanyu{border:solid 1px #5589AA;color: #000000 ; font-size: 12px;}
.font {  filter: DropShadow(Color=#cccccc, OffX=2, OffY=1, Positive=2); text-decoration: none; font-size: 9pt}
</style>
<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <tr> 
          <td bgcolor="#E8E8E8">��Ʒ��������</font></td>
		  <td bgcolor="#E8E8E8" colspan="5">&nbsp;</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">���</td>
          <td colspan="2" bgcolor="#F5F5F5" class="chinese">�����������</td>
          <td colspan="2" bgcolor="#F5F5F5" class="chinese">һ���������</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">����</td>
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
          
    <td bgcolor="#FFFFFF" class="chinese" align="center">&nbsp;</td>
          <td width="17%" align="center" bgcolor="#FFFFFF" class="chinese"><div align="center"><%=rs("p_small_type")%></div></td>
          
    <td width="18%" align="center" bgcolor="#FFFFFF" class="chinese"><%=rs("p_small_type_e")%></td>
          
    <td width="18%" align="center" bgcolor="#FFFFFF" class="chinese"><%=rs("p_type")%></td>
          
    <td width="17%" align="center" bgcolor="#FFFFFF" class="chinese"><%=rs("p_type_e")%></td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="admin_p_small.asp?id=<%=rs("p_small_id")%>&action=editp#editp">�޸�</a> 
            <a href="admin_p_small.asp?id=<%=rs("p_small_id")%>&action=delp#delp">ɾ��</a></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          
    <td bgcolor="#FFFFFF" colspan="6" class="chinese"><font color="#FF0000">&nbsp;</font></td>
        </tr>
<%End if
End if%>
</table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_p_small.asp">
        <tr>
          
      <td class="chinese" align="right"> /ҳ,����¼/ƪÿҳ. <a href="admin_p_small.asp?page=<%=page%>" title="��һҳ">>></a> 
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>
      <br>
	  <%if request.QueryString("action")="newp" then%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="newp">�µķ���</a></font></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">�������� :
              <input name="p_small_type" type="text" class="textarea" id="p_small_type4"  size="20">
            Ӣ�ķ��� :
            <input name="p_small_type_e" type="text" class="textarea" id="p_small_type_e"  size="20"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">���ڴ��� :              
              <select name="p_type" id="p_type">
                <option value="" selected>-��ѡ��һ������-</option>
                <% 
			'###############ȡ��������#################
			dim sql_news,rs_news
			sql_news="select p_type,p_type_e from p_class"
            set rs_news=server.createobject("adodb.recordset")
            rs_news.open sql_news,conn,1,1
			if Not(rs_news.BOF and rs_news.EOF) then 
			Do While Not rs_news.EOF 
			p_type_e=rs_news("p_type_e")
			  %>
                <option value="<%= rs_news("p_type") %>"><%= rs_news("p_type") %></option>
                <% 
			  rs_news.movenext
			  loop
			  rs_news.close
			  set rs_news=Nothing
			  'conn.close
			  'set conn=Nothing
			  End if
			'####################ȡ���������#####################  
			  %>
              </select></td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="newp">
		  <input type="hidden" name="p_type_e" value="<%=p_type_e%>">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
if request.QueryString("action")="editp" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>��ָ�������Ķ���"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>�Ƿ���ID������"
	Call diserror()
	response.End
  End if
End if
sql="select * from p_class_small where p_small_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="editp">�޸ķ���</a></font></td>
          </tr><tr>
            <td bgcolor="#FFFFFF" class="chinese">��������:
            <input name="p_small_type" type="text" class="textarea" id="p_small_type3" value="<%=rs("p_small_type")%>" size="20">
            ��������:
            <input name="p_small_type_e" type="text" class="textarea" id="p_small_type5" value="<%=rs("p_small_type_e")%>" size="20"></td>
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">���ڴ��� :
              <select name="p_type" id="p_type">
                <% 
			'###############ȡ��������#################
			dim sql_class,rs_class
            sql_class="select p_type,p_type_e from p_class"
            set rs_class=server.createobject("adodb.recordset")
            rs_class.open sql_class,conn,1,1
			if Not(rs_class.BOF and rs_class.EOF) then 
			Do While Not rs_class.EOF 
			p_type_e=rs_class("p_type_e")  
			  %>
                <option value="<%= rs_class("p_type") %>" <% if rs_class("p_type")=rs("p_type") then Response.Write("selected") End if %>><%= rs_class("p_type") %></option>
                <% 
			  rs_class.movenext
			  loop
			  rs_class.close
			  set rs_class=Nothing
			  'conn.close
			  'set conn=Nothing
			  End if
			'####################ȡ���������#####################  
			  %>
              </select></td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button"> 
              <input type="reset" name="Reset" value="�������" class="button">
              [<a href="admin_p.asp">����</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("p_small_id")%>">
		  <input type="hidden" name="p_type_e" value="<%=p_type_e%>">
		  <input type="hidden" name="action" value="editp">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
if request.QueryString("action")="delp" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>��ָ�������Ķ���"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>�Ƿ���ID������"
	Call diserror()
	response.End
  End if
End if
sql="select * from p_class_small where p_small_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="delp" id="delp">ɾ������</a>&nbsp;&nbsp;ע�⣺ɾ�����ཫͬʱɾ����������,������ʹ���޸Ĺ��ܣ�</font></td>
          </tr>
          <tr> 
            
      <td bgcolor="#FFFFFF" class="chinese">��������: [ ]�����ڴ���: [] </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��"> 
              [<a href="admin_p.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("p_small_id")%>">
          <input type="hidden" name="action" value="delp">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
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
End sub
%><marquee scrollAmount=10000 width="1" height="6">
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