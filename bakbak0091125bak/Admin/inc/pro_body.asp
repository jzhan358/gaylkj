<%sub adminpro_body()
dim totalnews,Currentpage,totalpages,i
sql="select * from p_info order by p_id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<HTML><HEAD><TITLE>��������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="inc/djcss.css" type=text/css rel=StyleSheet>
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
goaler[0] = new Array("�޷���","","");
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

function changelocation(location_id)//����һ�������ֵ,�Ӷ��ı��������
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
    <td bgcolor="#E8E8E8" colspan="3">��Ʒ����(ͼƬ�ϴ�ʱ��ע��ͼƬ�Ĵ�С�͸�ʽ��ֻ�����ϴ���С��200K���µġ�.jpg����.gif����β��ͼƬ��)</td>
  </tr>
  <tr bgcolor="#E8E8E8" align="center"> 
    <td class="chinese" width="10%" bgcolor="#F5F5F5">���</td>
    <td bgcolor="#F5F5F5" class="chinese">��Ʒ����</td>
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
    <td bgcolor="#FFFFFF" class="chinese" align="center"><%= rs("p_id") %>��</td>
    <td bgcolor="#FFFFFF" class="chinese"></span><%= rs("p_name") %>/<%= rs("p_name_e") %></td>
    <td bgcolor="#FFFFFF" class="chinese" align="center"><a href="admin_pro.asp?id=<%=rs("p_id")%>&action=editpro#editpro">�޸�</a> <a href="admin_pro.asp?id=<%=rs("p_id")%>&action=delpro#delpro">ɾ�� </td>
  </tr>
  <%i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then%>
  <tr align="center"> 
    <td bgcolor="#FFFFFF" colspan="3" class="chinese">��ǰû�в�Ʒ��</td>
  </tr>
  <%End if
End if%>
</table>
      <table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
	    <form name="form100" method="post" action="admin_pro.asp">
        <tr>
          
     <td class="chinese" align="right"> <%if currentpage>1 then%><a href="admin_pro.asp?page=<%=currentpage-1%>" title="��һҳ">��һҳ</a><%end if%>&nbsp;<%if currentpage<totalpages then%><a href="admin_pro.asp?page=<%=currentpage+1%>" title="��һҳ">��һҳ</a><%end if%>&nbsp;��ǰҳ/�ܹ�ҳ:<font color=red><%=currentpage%></font>/<font color=red><%=totalpages%></font>,��ǰ��¼/ÿҳ�ܹ���¼:<font color=red><%=i%></font> / <font color=red><%=newsperpage%></font>
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="text" name="page" class="textarea" size="4">
              <input type="submit" name="Submit" value="Go" class="button"></td>
        </tr>
		</form>
      </table>
      <br>
	  <%if request.QueryString("action")="newpro" then%>
      
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="">
    <INPUT name="upfile" onchange=return(OnUpFile()) type=hidden>
    <tr> 
      <td colspan="2" bgcolor="#E8E8E8"><a name="newpro">�µĲ�Ʒ</a></font></td>
    </tr>
    <tr> 
      <td width="36%" height="25" bgcolor="#FFFFFF" class="chinese">&nbsp;��Ʒ����: 
        <input name="p_name" type="text" class="textarea" id="p_name" size="30" maxlength="30"> 
        &nbsp;<span class="style1">*</span> &nbsp; </td>
      <td width="64%" bgcolor="#FFFFFF" class="chinese">Ӣ������: 
        <input name="p_name_e" type="text" class="textarea" id="p_name_e" size="30" maxlength="30"></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;��Ʒ�ͺ�: 
        <input name="p_spec" type="text" class="textarea" id="p_spec" value="1.03V" size="30" maxlength="30"> 
        &nbsp;<span class="style1">*</span> </td>
      <td bgcolor="#FFFFFF" class="chinese">Ӣ���ͺ�: 
        <input name="p_spec_e" type="text" class="textarea" id="p_spec_e" value="ydaer_1" size="30" maxlength="30"></td>
    </tr>
    <tr> 
      <td height="12" bgcolor="#FFFFFF" class="chinese">&nbsp;��Ʒ���: 
       <!--<select name="p_type" size="1" id="p_type" onChange="changelocation(document.form1.p_type.options[document.form1.p_type.selectedIndex].value)">-->
		<select name="p_type" size="1" id="p_type">
          <%set rs=conn.execute("select p_id,p_type from p_class")
          if rs.eof then%>
          <option selected value="">��һ������</option>
          <%else%>
          <option selected value="">һ������</option>
          <%do while not rs.eof
		  %>
          <option value="<%=rs("p_type")%>"><%=rs("p_type")%></option>
          <%rs.movenext
             loop
          end if%>
        </select> <!--<select name="p_small_type" >
          <option selected value="">��������</option>
        </select> --></td>
      <td height="12" bgcolor="#FFFFFF" class="chinese"><!--<input name="p_other2" type="checkbox" id="p_other2" value="1">
        ��Ϊ��Ա�ɲ鿴(ע�⣺��ѡΪ�����˶����Բ鿴)-->
        ������ֵ:
        <input name="p_order" type="text" class="textarea" id="p_order" value="0" size="15" maxlength="15">
���Բ���,��ֵԽ��,�ŵ�Խǰ(��ҳ����ǰ20����Ʒ��Ϣ)</td>
    </tr>
  <tr> 
      <td height="15" colspan="2" bgcolor="#FFFFFF" align="left">��Ʒ����ͼ:<input name="p_pic" type="text" class="textarea" id="p_pic"  value="" size="30" maxlength="100">
        �ϴ�ͼƬ��<iframe name=uploadformu frameborder=0 width=400 height=22 scrolling=no src=admin_upload.asp?id=p_pic></iframe><font color="#FF0000">(�ߴ磺160*110)</font></td>
    </tr> <tr> 
      <td height="15" colspan="2" bgcolor="#FFFFFF" align="left">��Ʒ��ͼ:<input name="p_pic2" type="text" class="textarea" id="p_pic2"  value="" size="30" maxlength="100">
        �ϴ�ͼƬ��<iframe name=uploadformu frameborder=0 width=400 height=22 scrolling=no src=admin_upload.asp?id=p_pic2></iframe><font color="#FF0000">(�ߴ�:500*500����)</font> </td>
    </tr> 
    <tr> 
      <td height="25" colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">&nbsp;���ļ�飺 
        ���±༭��ʹ�ú�Word����,����ע�����ݲ�Ҫ̫��������1500����Ϊ�ѡ�</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF"><textarea name="p_jianjie" style="display:none"></textarea><iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=p_jianjie&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF"> Ӣ�ļ�飺</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF"><textarea name="p_jianjie_e" style="display:none"></textarea><iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=p_jianjie_e&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
        </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" align="center" bgcolor="#F5F5F5" class="chinese"> 
        <input type="submit" name="Submit2" value="ȷ������" class="button"> &nbsp; 
        <input type="reset" name="Reset" value="�������" class="button"> </td>
    </tr>
    <input type="hidden" name="action" value="newpro">
    <input type="hidden" name="MM_insert" value="true">
  </form>
</table>
	  <%End if
if request.QueryString("action")="editpro" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>��ָ�������Ķ���"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>�Ƿ��Ĳ�ƷID������"
	Call diserror()
	response.End
  End if
End if
sql="select * from p_info where p_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      
<table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
  <form name="form1" method="post" action="">
    <INPUT name="upfile" onchange=return(OnUpFile()) type=hidden>
    <input name="upfile2" type="hidden" value="<%= rs("p_pic") %>">
    <tr> 
      <td colspan="2" bgcolor="#E8E8E8"><a name="newpro">�޸Ĳ�Ʒ</a></font>&nbsp; 
        (ע�⣺��Ҫ�޸ĵĵط����벻Ҫ�Ķ�����ԭ������</td>
    </tr>
    <tr> 
      <td width="38%" bgcolor="#FFFFFF" class="chinese">&nbsp;��Ʒ����: 
        <input name="p_name" type="text" class="textarea" id="p_name" value="<%= rs("p_name") %>" size="30" maxlength="30"> 
        &nbsp;<span class="style1">*</span>&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; </td>
      <td width="62%" bgcolor="#FFFFFF" class="chinese">Ӣ������: 
        <input name="p_name_e" type="text" class="textarea" id="p_name_e" value="<%= rs("p_name_e") %>" size="30" maxlength="30"></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;��Ʒ�ͺ�: 
        <input name="p_spec" type="text" class="textarea" id="p_spec" value="<%= rs("p_spec") %>" size="30" maxlength="30"> 
        &nbsp;<span class="style1">*</span> </td>
      <td bgcolor="#FFFFFF" class="chinese">Ӣ���ͺ�: 
        <input name="p_spec_e" type="text" class="textarea" id="p_spec_e" value="<%= rs("p_spec_e") %>" size="30" maxlength="30"></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="chinese">&nbsp;��Ʒ���: 
        <select name="p_type" id="p_type">
          <% 
			'###############ȡ��������#################
			'dim sql_p,rs_p
			sql_p="select p_id,p_type from p_class"
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
			Do While Not rs_p.EOF

			  %>
          <option value="<%= rs_p("p_type") %>" <% if rs_p("p_type")=rs("p_type") then Response.Write("selected") End if %> ><%= rs_p("p_type") %></option>
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
        </select> <!--<select name="p_small_type">
          <option selected value="<%'=rs("p_small_type")%>"><%'=rs("p_small_type")%></option>
        </select>--> </td>
     <td height="30" bgcolor="#FFFFFF" class="chinese">������ֵ: 
        <input name="p_order" type="text" class="textarea" id="p_order" value="<%= rs("p_order") %>" size="15" maxlength="15">
        ���Բ���,��ֵԽ��,�ŵ�Խǰ(��ҳ����ǰ20����Ʒ��Ϣ)</td>
    </tr>
	 <tr> 
      <td height="15" colspan="2" bgcolor="#FFFFFF" align="left">��Ʒ����ͼ:<input name="p_pic" type="text" class="textarea" id="p_pic"  value="<%= rs("p_pic") %>" size="30" maxlength="100">
        �ϴ�ͼƬ��<iframe name=uploadformu frameborder=0 width=400 height=22 scrolling=no src=admin_upload.asp?id=p_pic></iframe><font color="#FF0000">(�ߴ磺160*110)</font></td>
    </tr> <tr> 
      <td height="15" colspan="2" bgcolor="#FFFFFF" align="left">��Ʒ��ͼ:<input name="p_pic2" type="text" class="textarea" id="p_pic2"  value="<%= rs("p_pic2") %>" size="30" maxlength="100">
        �ϴ�ͼƬ��<iframe name=uploadformu frameborder=0 width=400 height=22 scrolling=no src=admin_upload.asp?id=p_pic2></iframe><font color="#FF0000">(�ߴ�:500*500����)</font> </td>
    </tr> 
    
    <tr>
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">&nbsp;���ļ�飺</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">
        <textarea name="p_jianjie" style="display:none"></textarea><iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=p_jianjie&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe> 
      </td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese"> Ӣ�ļ�飺</td>
    </tr>
    <tr> 
      <td colspan="2" align="left" bgcolor="#FFFFFF" class="chinese">
        <textarea name="p_jianjie_e" style="display:none"></textarea><iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=p_jianjie_e&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" align="center" bgcolor="#F5F5F5" class="chinese"> 
        <input type="submit" name="Submit2" value="ȷ���޸�" class="button"> &nbsp; 
        <input type="reset" name="Reset" value="�������" class="button"> </td>
    </tr>
    <input type="hidden" name="id" value="<%= cint(request.querystring("id")) %>">
    <input type="hidden" name="action" value="editpro">
    <input type="hidden" name="MM_insert" value="true">
  </form>
</table>
	  <%End if
if request.QueryString("action")="delpro" then
if request.querystring("id")="" then
  errmsg=errmsg+"<br>"+"<li>��ָ�������Ķ���"
  Call diserror()
  response.End
else
  if Not isinteger(request.querystring("id")) then
    errmsg=errmsg+"<br>"+"<li>�Ƿ��Ĳ�ƷID������"
	Call diserror()
	response.End
  End if
End if
sql="select * from p_info where p_id="&cint(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="newpro">ɾ����Ʒ</a></font></td>
          </tr>
		   <tr> 
            <td align="center" valign="top" bgcolor="#FFFFFF" class="chinese">&nbsp;<span class="style1">�� ȷ �� Ҫ ɾ �� ��</span></td>
          </tr>
         
          
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center">
              <input type="submit" name="Submit2" value="ȷ��ɾ��" class="button">
&nbsp; ��<a href="admin_pro.asp">����</a>�� </td>
          </tr>
		  <input type="hidden" name="id" value="<%= request.querystring("id") %>">
		  <input type="hidden" name="action" value="delpro">
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
