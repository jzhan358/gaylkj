<%sub p_class_body()
dim totalp,Currentpage,totalpages,i
sql="select * from p_class order by p_id ASC"
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
          <td width="12%" bgcolor="#E8E8E8"��width="10%">��Ŀ�������</font></td>
          <td width="18%" bgcolor="#FFFFFF"��width="10%"><div align="center"><a href="admin_p_class.asp?action=new" class="style1"> ���һ������</a></div></td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td colspan="2" bgcolor="#F5F5F5" class="chinese">����Ŀ��������</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">�١�����ѡ����</td>
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
          
    <td colspan="2" bgcolor="#F5F5F0" class="chinese">������<img src="images/tree_1.gif" width="15" height="15">��&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000"><strong>��Ѱ汾�޴˹��ܣ�</strong></font></td>
          <td align="center" bgcolor="#F5F5F0" class="chinese">
		  <a href="admin_p_class.asp?p_type=<%=rs("p_type")%>&action=p#p">��Ӷ�������</a> |
		  <a href="admin_p_class.asp?id=<%=rs("p_id")%>&action=edit#edit">�޸�</a> |
          <a href="admin_p_class.asp?id=<%=rs("p_id")%>&action=del#del">ɾ��</a></td>
        </tr>
      <%
	  set rsSmallClass=server.CreateObject("adodb.recordset")
	  rsSmallClass.open "Select * From p_class_small Where p_type='" & rs("p_type") & "'",conn,1,1
	  if not(rsSmallClass.bof and rsSmallClass.eof) then
	  Do While Not rsSmallClass.eof
	  %>
		<tr>
          
    <td colspan="2" bgcolor="#FFFFFF" class="chinese">��������<img src="images/tree_2.gif" width="15" height="15">��</td>
          <td align="center" bgcolor="#FFFFFF" class="chinese">
		  ��������������<a href="admin_p_class.asp?id=<%=rsSmallClass("p_small_id")%>&action=edits#edits">�޸�</a> |
          <a href="admin_p_class.asp?id=<%=rsSmallClass("p_small_id")%>&action=dels#dels">ɾ��</a></td>
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
          
    <td bgcolor="#FFFFFF" colspan="3" class="chinese"><font color="#FF0000"><strong>��Ѱ汾�޴˹��ܣ�</strong></font></td>
        </tr>
<%End if
End if%>
</table>
      <table align="center" width="80%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_p_class.asp">
        <tr>
          
      <td class="chinese" align="right"> /ҳ,����¼/ƪÿҳ. <a href="admin_p_class.asp?page=<%=page%>" title="��һҳ">>></a> 
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
            <td bgcolor="#E8E8E8"><a name="add" id="add">�����Ŀ����</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">�������� :
        <input name="p_type" type="text" class="textarea"  size="20">
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="new">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
if request.QueryString("action")="edit" then
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
sql="select * from p_class where p_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="edit" id="edit">�޸���Ŀ����</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">��������: 
              <input type="text" name="p_type" size="20" class="textarea" value="<%=rs("p_type")%>"> 
            </td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button"> 
              <input type="reset" name="Reset" value="�������" class="button">
              [<a href="admin_p_class.asp">����</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("p_id")%>">
		  <input type="hidden" name="action" value="edit">
          <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
if request.QueryString("action")="del" then
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
sql="select * from p_class where p_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="del" id="del">ɾ����Ŀ����</a>&nbsp;&nbsp;ע�⣺ɾ�����ཫͬʱɾ����������,������ʹ���޸Ĺ��ܣ�</font></td>
          </tr>
          <tr> 
            
      <td bgcolor="#FFFFFF" class="chinese">��������: </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��"> 
              [<a href="admin_p_class.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("p_id")%>">
          <input type="hidden" name="action" value="del">
          <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if%>
      	
       <%'���������Զ������������ӣ��޸ģ�ɾ�����������
	   if request.QueryString("action")="p" then%>
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="adds" id="adds">��Ӷ�������</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">�������ࣺ
			<select name="p_type" id="p_type">
				<% 
			'###############ȡ��������#################
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
			'####################ȡ���������#####################  
			  %>
              </select>
			</td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">�������� :
            <input name="p_small_type" type="text" class="textarea" id="p_small_type"  size="20"></td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="p">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
if request.QueryString("action")="edits" then
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
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="edits" id="edits">�޸Ķ�������</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">��������:
              
              <select name="p_type" id="p_type">
                <% 
			'###############ȡ��������#################
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
			'####################ȡ���������#####################  
			  %>
              </select></td>
          <tr>
            <td bgcolor="#FFFFFF" class="chinese">��������:
            <input name="p_small_type" type="text" class="textarea" id="p_small_type" value="<%=rs("p_small_type")%>" size="20"></td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button"> 
              <input type="reset" name="Reset" value="�������" class="button">
              [<a href="admin_p_class.asp">����</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("p_small_id")%>">
		  <input type="hidden" name="action" value="edits">
          <input type="hidden" name="MM_insert" value="true">
        </form>
</table>
	  <%End if
if request.QueryString("action")="dels" then
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
      <table align="center" width="80%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="dels" id="dels">ɾ����������</a>&nbsp;&nbsp;ע�⣺ɾ�����ཫͬʱɾ����������,������ʹ���޸Ĺ��ܣ�</font></td>
          </tr>
          <tr> 
            
      <td bgcolor="#FFFFFF" class="chinese">��������:����������: </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��"> 
              [<a href="admin_p_class.asp">����</a>] </td>
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
