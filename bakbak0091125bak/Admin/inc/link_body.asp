<%sub adminlink_body()
dim totalnews,Currentpage,totalpages,i
sql="select * from link_class order by l_id ASC"
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
          <td bgcolor="#E8E8E8">��ϵ����</font></td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">���</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">��������</td>
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
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("l_id")%></td>
          <td bgcolor="#FFFFFF" class="chinese"><%=rs("l_type")%>&nbsp;&nbsp;&nbsp;&nbsp;<% Call v1(rs("l_id"),rs("l_type")) %></td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="admin_link.asp?id=<%=rs("l_id")%>&action=editlink#editlink">�޸�</a> 
            <a href="admin_link.asp?id=<%=rs("l_id")%>&action=dellink#dellink">ɾ��</a></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">��ǰû�в��ţ�</td>
        </tr>
<%End if
End if%>
</table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_link.asp">
        <tr>
          <td class="chinese" align="right"><%=currentpage%> /<%=totalpages%>ҳ,<%=totalnews%>����¼/<%=newsperpage%>ƪÿҳ. 
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
              <a href="admin_link.asp?page=<%=i%>"><%=i%></a> 
              <%End if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
End if%>
              <a href="admin_link.asp?page=<%=page%>" title="��һҳ">>></a> 
              <%End if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>
      <br>
	  <%if request.QueryString("action")="newlink" then%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="newlink">�µĲ���</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">�������� :
        <input name="l_type" type="text" class="textarea"  size="20">
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="newlink">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
if request.QueryString("action")="editlink" then
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
sql="select * from link_class where l_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="editlink">�޸�����</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">��������: 
              <input type="text" name="l_type" size="20" class="textarea" value="<%=rs("l_type")%>"> 
            </td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button"> 
              <input type="reset" name="Reset" value="�������" class="button">
              [<a href="admin_link.asp">����</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("l_id")%>">
		  <input type="hidden" name="action" value="editlink">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
if request.QueryString("action")="dellink" then
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
sql="select * from link_class where l_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="dellink" id="dellink">ɾ������</a>&nbsp;&nbsp;ע�⣺ɾ�����Ž�ͬʱɾ����������,������ʹ���޸Ĺ��ܣ�</font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">��������: <%=rs("l_type")%> 
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��"> 
              [<a href="admin_link.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("l_id")%>">
          <input type="hidden" name="action" value="dellink">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
	  
	  if request.QueryString("action")="viewlink" then
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
sql="select * from link_info where link_class_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form8" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="viewlink" id="viewlink">�鿴����</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><% = request.querystring("l_type")%>
            </td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="textarea">&nbsp;�� ϵ �ˣ�
            <input name="link_name" type="text" value="<%=rs("link_name")%>" size="20" maxlength="20"></td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="textarea">&nbsp;�绰���룺            
            <input name="link_tel" type="text" value="<%=rs("link_tel")%>" size="20" maxlength="20"></td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="textarea">&nbsp;������룺
            <input name="link_fax" type="text" value="<%=rs("link_fax")%>" size="20" maxlength="20"></td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="textarea">&nbsp;�������䣺
          <input name="link_email" type="text" value="<%=rs("link_email")%>" size="20" maxlength="20"></td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
          <input name="Submit" type="submit" class="button" id="Submit" value="ȷ���޸�"> 
              [<a href="admin_link.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=cint(request.querystring("id"))%>">
          <input type="hidden" name="action" value="edittextlink">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
	  
	  
	  if request.QueryString("action")="addlink" then
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
     %>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form9" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="addlink" id="addlink">��������</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><% = request.querystring("l_type")%>
            </td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="textarea">&nbsp;�� ϵ �ˣ�
            <input name="link_name" type="text" size="20" maxlength="20"></td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="textarea">&nbsp;�绰���룺
            <input name="link_tel" type="text" size="20" maxlength="20"></td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="textarea">&nbsp;������룺
            <input name="link_fax" type="text" size="20" maxlength="20"></td>
          </tr>
		  <tr> 
            <td bgcolor="#FFFFFF" class="textarea">&nbsp;�������䣺
            <input name="link_email" type="text" size="20" maxlength="20"></td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ������"> 
              [<a href="admin_link.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=cint(request.querystring("id"))%>">
          <input type="hidden" name="action"    value="addlink">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<% End if %>
      <br>
    </td>
  </tr>
  <tr> 
    <td colspan="3" height="1" background="images/dotlineh.gif"></td>
  </tr>
</table>
<% 
       sql_c="select * from config"
       set rs_c=server.createobject("adodb.recordset")
       rs_c.open sql_c,conn,1,1
     %>
      <table width="98%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#C0C0C0" >
        <% if rs_c.BOF and rs_c.EOF then %>
		<tr>
          <td height="30" colspan="4">&nbsp;û�����빫˾������Ϣ&nbsp;&nbsp; ��<a href="admin_config.asp?action=addconfig">���������Ϣ</a>��</td>
        </tr>
		<% 
		elseif Not(rs_c.EOF or rs_c.BOF) then %>
		<tr bgcolor="#E8E8E8">
          <td height="30" colspan="4">&nbsp;��˾������Ϣ�鿴</td>
        </tr>
        <tr>
          <td width="13%" height="30" align="right" valign="middle">��˾���ƣ�</td>
          <td colspan="3" align="left" valign="middle"><%= rs_c("c_incname") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">��˾��ַ��</td>
          <td width="41%" height="30" align="left" valign="middle"><%= rs_c("c_addr") %></td>
          <td width="14%" align="right" valign="middle">��˾��ַ��</td>
          <td width="32%" align="left" valign="middle"><%= rs_c("c_web") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">�绰���룺 </td>
          <td height="30" align="left" valign="middle"><%= rs_c("c_tel") %></td>
          <td align="right" valign="middle">�������䣺</td>
          <td align="left" valign="middle"><%= rs_c("c_email") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">������룺</td>
          <td height="30" align="left" valign="middle"><%= rs_c("c_fax") %></td>
          <td align="right" valign="middle">�� ϵ �ˣ�</td>
          <td align="left" valign="middle"><%= rs_c("c_linkman") %></td>
        </tr>
		<tr align="center" valign="middle">
          <td height="30" colspan="4">��<a href="admin_config.asp?action=editconfig">�޸���Ϣ</a>��</td>
        </tr>
		<% 
		End if
		rs_c.close
		set rs_c=Nothing
		 %>
      </table>
<%
rs.close
set rs=Nothing

End sub
sub v1(id,l_type)
v_sql="select * from link_info where link_class_id="&id
set v_rs=server.createobject("adodb.recordset")
v_rs.open v_sql,conn,1,1
if v_rs.EOF and v_rs.BOF then
%>
<a href="admin_link.asp?id=<%=id%>&l_type=<%=l_type%>&action=addlink">[���Ӳ�������]</a>
<%
else
%>
<a href="admin_link.asp?id=<%=id%>&l_type=<%=l_type%>&action=viewlink">[�޸Ĳ�������]</a>
&nbsp;&nbsp;
<a href="../showlink.asp">[�鿴��������]</a>

<%
End if
v_rs.close
set v_rs=Nothing
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