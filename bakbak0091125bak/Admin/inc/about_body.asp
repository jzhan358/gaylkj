<%sub adminabout_body()
dim totalnews,Currentpage,totalpages,Language
Language=request.QueryString("language")
if Language<>"" then
sql="select * from inc_class  where i_language=" & language &" order by i_id ASC"
else
sql="select * from inc_class  where i_language=0 order by i_id ASC"
end if
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
          <td bgcolor="#E8E8E8">��������</font></td>
		  <td bgcolor="#FFFFFF" colspan="2">[<a href="admin_about.asp?language=0">������Ŀ</a>][<a href="admin_about.asp?language=1">Ӣ����Ŀ</a>]</td>
        </tr>
        <tr bgcolor="#E8E8E8" align="center"> 
          <td class="chinese" width="10%" bgcolor="#F5F5F5">���</td>
          <td class="chinese" width="70%" bgcolor="#F5F5F5">��Ŀ����</td>
          <td class="chinese" width="20%" bgcolor="#F5F5F5">����</td>
        </tr>
<%if Not rs.EOF then
dim newsperpage
newsperpage=10
rs.movefirst
rs.pagesize=newsperpage
if trim(request("page"))<>"" then
   currentpage=clng(request("page"))
   if currentpage>rs.pagecount then
      currentpage=rs.pagecount
   End if
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
          <td bgcolor="#FFFFFF" class="chinese" align="center"><%=rs("i_id")%></td>
          <td bgcolor="#FFFFFF" class="chinese"><%=rs("i_type")%></td>
          <td bgcolor="#FFFFFF" class="chinese" align="center"><!--<a href="admin_about.asp?id=<%'=rs("i_id")%>&Language=<%'=rs("i_language")%>&action=editabout#editabout">�޸�</a> 
            <a href="admin_about.asp?id=<%'=rs("i_id")%>&Language=<%'=rs("i_language")%>&action=delabout#delabout">ɾ��</a>--><% Call v1(rs("i_id"),rs("i_type")) %></td>
        </tr>
<%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
        <tr align="center"> 
          <td bgcolor="#FFFFFF" colspan="3" class="chinese">��ǰû����Ŀ��</td>
        </tr>
<%End if
End if%>
</table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form2" method="post" action="admin_about.asp">
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
              <a href="admin_about.asp?page=<%=i%>"><%=i%></a> 
              <%End if
next
if totalpages>currentpage then
if request("page")="" then
page=1
else
page=request("page")+1
End if%>
              <a href="admin_about.asp?page=<%=page%>" title="��һҳ">>></a> 
              <%End if%>
              &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>
      <br>
	  <%if request.QueryString("action")="newabout" then%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="newabout">�µ���Ŀ</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">��Ŀ���� :
        <input name="i_type" type="text" class="textarea"  size="20">
        �汾��
        <input name="i_language" type="radio" value="0" checked>
        ����
        <input type="radio" name="i_language" value="1">
        Ӣ��
        
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="ȷ������" class="button">
              <input type="reset" name="Reset" value="�������" class="button">
            </td>
          </tr>
		  <input type="hidden" name="action" value="newabout">
		  <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
if request.QueryString("action")="editabout" then
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
sql="select * from inc_class where i_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="editabout">�޸���Ŀ</a></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">��Ŀ����: 
              <input type="text" name="i_type" size="20" class="textarea" value="<%=rs("i_type")%>">
              �汾��
              
			  <input type="radio" name="i_language" value="0" <%if rs("i_language")="0" then%> checked<%end if%>>
����
<input type="radio" name="i_language" value="1" <%if rs("i_language")="1" then%> checked<%end if%>>
Ӣ��            </td>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input type="submit" name="Submit" value="ȷ���޸�" class="button"> 
              <input type="reset" name="Reset" value="�������" class="button">
              [<a href="admin_about.asp">����</a>] </td>
          </tr>
		  <input type="hidden" name="id" value="<%=rs("i_id")%>">
		  <input type="hidden" name="action" value="editabout">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
if request.QueryString("action")="delabout" then
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
sql="select * from inc_class where i_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="delabout" id="delabout">ɾ����Ŀ</a>&nbsp;&nbsp;ע�⣺ɾ����Ŀ��ͬʱɾ����Ŀ����,������ʹ���޸Ĺ��ܣ�</font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">��Ŀ����: <%=rs("i_type")%> 
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ��ɾ��"> 
              [<a href="admin_about.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=rs("i_id")%>">
          <input type="hidden" name="action" value="delabout">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
	  
	  if request.QueryString("action")="viewabout" then
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
sql="select * from inc_info where inc_class_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form8" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="viewabout" id="viewabout">�鿴����</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><% = request.querystring("i_type")%>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
              <textarea name="inc_text" style="display:none"><%=Server.HTMLEncode(rs("inc_text"))%></textarea>
             <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=inc_text&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ���޸�"> 
              [<a href="admin_about.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=cint(request.querystring("id"))%>">
		  <input type="hidden" name="i_language" value="<% = request.querystring("i_language")%>">
          <input type="hidden" name="action" value="edittextabout">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if
	  
	  
	  if request.QueryString("action")="addabout" then
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
            <td bgcolor="#E8E8E8"><a name="addabout" id="addabout">��������</a></font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><% = request.querystring("i_type")%>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
              <textarea name="inc_text" style="display:none"></textarea>
             <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=inc_text&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="ȷ������"> 
              [<a href="admin_about.asp">����</a>] </td>
          </tr>
          <input type="hidden" name="id" value="<%=cint(request.querystring("id"))%>">
		  <input type="hidden" name="i_language" value="<% = request.querystring("i_language")%>">
          <input type="hidden" name="action" value="addabout">
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
rs.close
set rs=Nothing

End sub
sub v1(id,i_type)
v_sql="select * from inc_info where inc_class_id="&id
set v_rs=server.createobject("adodb.recordset")
v_rs.open v_sql,conn,1,1
if v_rs.EOF and v_rs.BOF then
%>
<a href="admin_about.asp?id=<%=id%>&i_type=<%=i_type%>&action=addabout">[������Ŀ����]</a>
<%
else
%>
<a href="admin_about.asp?id=<%=id%>&i_type=<%=i_type%>&action=viewabout">[�޸���Ŀ����]</a>

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