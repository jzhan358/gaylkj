<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="config.asp"-->
<!--#include file="eintop.asp" -->
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
<TABLE cellSpacing=0 cellPadding=0 width=778 bgColor=#ffffff border=0 height="310">
  <TR>
    <TD vAlign=top height="310">
      <TABLE cellSpacing=0 cellPadding=0 width=9 border=0>
        <TBODY>
        <TR>
          <TD>��</TD></TR></TBODY></TABLE></TD>
    <TD width="100%" height="310">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="321">
        <TBODY>
        <TR>
          <TD height="321">
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="318">
              <TBODY>
              <TR>
                <TD vAlign=top width=61 height="318">
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG height=29 src="index/Top_Bn_Left.gif" 
                        width=190 border=0></TD></TR>
                    <TR>
                      <TD height=20><!-- //�����˵� -->
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD height="22"></TD></TR></TBODY></TABLE>
                                  <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                                    <TBODY>
                                      <TR> 
                                        <TD><img border="0" src="images/eLeft_About.gif"></TD>
                                      </TR>
                                      <tr> 
                                        <TD bgColor=#0BBAFD height=2></TD>
                                      </tr>
                                      <tr> 
                                        <TD height="15"> <TABLE borderColor=#111111 cellSpacing=0 cellPadding=2 width="100%" border=0 class="font22">
              <tr>
                        <TD width=100% height="8"> 
</td>
              </tr>
              <tr> 
                        <TD width=100% height="16"> 
<script language="JavaScript" type="text/JavaScript">
		             <!--
            function MM_jumpMenu(targ,selObj,restore){ //v3.0
                      eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
                      if (restore) selObj.selectedIndex=0;
                     }
            //-->
<!--
	function ToggleDisplay(oButton, oItems){
		if ((oItems.style.display == "") || (oItems.style.display == "block"))	{
			oItems.style.display = "none";
			oButton.src = "/images/tree_plus.gif";
		}	
		else {
			oItems.style.display = "block";
			oButton.src = "/images/tree_minus.gif";
		}
		return false;
	}
	function UnFoldAll() {
		temp = document.all.length;
	    for(i=0;i<temp;i++)
		    if(document.all[i].id.substr(0,3) =="Img")
			    document.all[i].src = "/images/tree_minus.gif"
	        else if(document.all[i].id.substr(0,4) =="Tree")
		        document.all[i].style.display ="block"
	}
	function FoldAll() {
		temp = document.all.length;
	    for(i=0;i<temp;i++)
		    if(document.all[i].id.substr(0,3) =="Img")
			    document.all[i].src = "/images/tree_plus.gif"
	        else if(document.all[i].id.substr(0,4) =="Tree")
		        document.all[i].style.display ="none"
	}
//-->
</SCRIPT>

<%   
sql = "SELECT * FROM P_class ORDER BY p_id"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.OPEN sql,Conn,0,1
if rs.eof and rs.bof then
Response.Write"no product type"
else
do while not rs.eof

sql = "SELECT * FROM P_class_small where p_type ='"&rs("p_type_e")&"'"
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.OPEN sql,Conn,1,1
if rs2.recordcount > 0 then 
Response.Write"&nbsp;<A border=0 style='CURSOR: hand' id=Img"&rs("p_id")&" onclick=ToggleDisplay(Img"&rs("p_id")&",Tree"&rs("p_id")&")><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type_e")&"</b></a>"
else
Response.Write"&nbsp;<A href=eProduct.asp?classtype="&rs("p_type_e")&" class='six'><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type_e")&"</b></a>"
end if
counts=conn.execute("Select count(*) from P_info where p_other = 0 and p_type='"&rs("p_type_e")&"'")(0)
Response.Write"(<FONT COLOR='800000'>"&counts&"</FONT>)<br>"
%>
</td></tr>
<tr>
  <td><img border="0" src="index/line.gif"></td>
</tr>
<tr><td>
                          <div id=Tree<%=rs("p_id")%>> 
                            <%
do while not rs2.eof
%>
                            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="100%">&nbsp; ��&nbsp;<a class="number" href='eProduct.asp?classtype=<%=rs2("p_type_e")%>&smallclass=<%=rs2("p_small_type_e")%>'><%=rs2("p_small_type_e")%></a></td>    
                              </tr>
                            </table>
                            <%
rs2.movenext
loop
rs2.close
Set rs2=Nothing
%>
                            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="100%" height="5"></td>
                              </tr>
                            </table>
                          </div>
                          <%
rs.movenext
loop
end if
rs.close
Set rs=Nothing
%>
<SCRIPT language=javascript>FoldAll();</SCRIPT>
                        </td>
                      </tr>
                    </table></TD>
                                      </tr>
                                      <TR> 
                                        <TD height=10></TD>
                                      </TR>
                                      <TR> 
                                        <TD height=3><img border="0" src="images/2Left_About2.gif"></TD>
                                      </TR>
                                      <tr> 
                                        <TD bgColor=#0BBAFD height=2></TD>
                                      </tr>
                                      <TR> 
                                        <TD> <DIV align=right> 
                                            <TABLE cellSpacing=0 cellPadding=0 width=175 
                              border=0>
                                              <TBODY>
                                                <TR> 
                                                  <TD height=5></TD>
                                                </TR>
                                                <tr> 
                                                  <TD align=middle width="163" height=60> 
                                                    <DIV align=center>
                                                      <table width="100%" height="34" border="0" cellpadding="0" cellspacing="0">
                                                        <form name="form1" method="post" action="Demo_eProduct.asp">
                                                            <td width="124" align="left"><input name="key" type="text" class="bd5" id="key" size="15"></td>
                                                            <td width="48" align="left"><input name="imageField" type="image" src="images/esubmit.gif" border="0"></td>
                                                        </form>
                                                      </table>
                                                    </DIV></TD>
                                                </tr>
                                              </TBODY>
                                            </TABLE>
                                          </DIV></TD>
                                      </TR>
                                    </TBODY>
                                  </TABLE></TD></TR></TBODY></TABLE></TD>
                <TD vAlign=top background=images/Jg_Left_Bg.gif height="318"><IMG 
                  src="images/Jg_Left.gif" border=0 width="14" height="218"></TD>
                <TD vAlign=top width="100%" height=318>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG src="images/Top_Bn_Foot.gif" border=0></TD></TR>
                    <TR>
                      <TD>
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD>
                              <TABLE cellSpacing=0 cellPadding=0 width=10 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD>��</TD></TR></TBODY></TABLE></TD>
                            <TD width="100%">
                              <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              border=0>
                                <TBODY>
                                <tr>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                <TBODY>
                                <TR>
                                <TD><IMG height=28 
                                src="images/Menu_Bg_Left.gif" width=15 
                                border=0></TD>
                                <TD width="100%" 
                                background=images/Menu_Bg_m.gif>
                                <P align=center><FONT 
                                color=#ffffff><a href="index.asp"><FONT 
                                color=#ffffff>Home</FONT></a>&nbsp; |&nbsp; <a href="eindex.asp"><FONT                                     
                                color=#ffffff>Company</FONT></a>&nbsp; |&nbsp; <a href="eNews.asp"><FONT    
                                color=#ffffff>News</FONT></a>&nbsp;                                        
                                |&nbsp; <a href="eProduct.asp"><FONT                                      
                                color=#ffffff>Products</FONT></a>&nbsp; |&nbsp; <a style="COLOR: #ffffff" href="demo_exs.asp">Sale</a>&nbsp;    
                                |&nbsp; <a href="eGuest.asp"><FONT                                     
                                color=#ffffff>Messenger</FONT></a> |&nbsp; <a style="COLOR: #ffffff" href="eOrder.asp">Order</a>&nbsp;    
                                |&nbsp; <a href="Demo_Guest.asp"><FONT                                     
                                color=#ffffff></FONT></a><a href="elinkus.asp"><FONT 
                                color=#ffffff>Contact</FONT></a></FONT></P></TD>
                                <TD><IMG height=28 
                                src="images/Menu_Bg_Right.gif" width=15 
                                border=0></TD></TR></TBODY></TABLE></TD>
                                </tr>
                                <TR>
                                <TD height=10></TD></TR>
                                <TR>
                                                <TD height=25> 
                                                  <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                                    <TBODY>
                                                      <tr>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                <TBODY>
                                <TR>
                                <TD width=43><img border="0" src="index/Products.gif"></TD>
                                <TD width="100%">
                                <P align=right>��</P></TD></TR></TBODY></TABLE></TD>
                                                      </tr>
                                                      <tr>
                                <TD height=5></TD>
                                                      </tr>
                                                      <tr>
                                <TD bgColor=#eaeaea height=1></TD>
                                                      </tr>
                                                    </TBODY>
                                                  </TABLE>
                                                        <%
set rs=server.createobject("ADODB.Recordset")
sql="select * from P_info where p_other = 0 "
classType = request("classType")
smallclass = request("smallclass")
key = request("key")
if classType <> "" then
	sql = sql & " and p_type_e = '"&classType&"'"
end if
if smallclass <> "" then
	sql = sql & " and p_small_type_e = '"&smallclass&"'"
end if
if key <> "" then
	sql = sql & " and (p_name_e like '%"&key&"%' or  p_spec_e like '%"&key&"%')"
end if

sql = sql & " order by p_id desc"
rs.open sql,conn,1,1

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

%>

                                                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                                    <tr>
                                  <TD height=5></TD>
                                                    </tr>
                                                    <tr> 
                                                      <td valign="top">
														The display��Products show - ->Total <%=rs.recordcount%> pcs, Total <%=rs.pagecount%>  page,This is the <%=pagecount%> page, 12 pcs of products per page      
                                                      </td>
                                                    </tr>
                                                    <tr> 
                                                      <td align="left" valign="top"> 
                                                        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="table3">
                                                          <%
						sql_type="select * from P_class_small where p_type = '" & classType & "' order by p_small_id desc"
						set rs_type=server.createobject("adodb.recordset")
						rs_type.open sql_type,conn,1,3
					%>
                                                          <tr> 
                                                            <%
					if rs_type.eof and rs_type.bof then 
					 else
							for i=1 to rs_type.recordcount
							 if rs_type.EOF then exit for 
					%>
                                                            <td height="25" align="center" class="line"><a href="eProduct.asp?classtype=<%=rs_type("p_type_e")%>&smallclass=<%=rs_type("p_small_type_e")%>" class="six"><%=rs_type("p_small_type_e")%></a></td>
                                                            <% if i mod 6 = 0 then %>
                                                          </tr>
                                                          <tr> 
                                                            <%end if%>
                                                            <%
						 rs_type.movenext
						 next
						 end if 
						 rs_type.close
					 %>
                                                        </table>
                                                        <table width="539" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="table3">
                                                          <tr> 
                                                            <%
					if rs.eof or rs.bof then
					  response.write("There are no products at present")
					 else
						rs.pagesize=12
						%>
                                                            <td align="center" class="line" width="539">  <%
						rs.AbsolutePage=pagecount

						for i=1 to rs.recordcount
							if rs.EOF then exit for 
					%>
                                                              <table width="131" border="0" cellspacing="0" cellpadding="0">
                                                                <tr> 
                                                                  <td height="88" align="center" valign="bottom"><a href="eshowProDetail.asp?ProID=<%=rs("p_id")%>"><img src="<%=rs("p_pic")%>" width="110" height="150" border="0"></a></td>
                                                                </tr>
                                                                <tr> 
                                                                  <td height="50" align="center" class="font9"> 
                                                                    Name��<A href="eshowProDetail.asp?ProID=<%=rs("p_id")%>" class="money"><%=rs("p_name_e")%></A><br>
                                                                    Item��<%=rs("p_spec_e")%> 
                                                                  </td>
                                                                </tr>
                                                              </table></td>
                                                            <% if i mod 4 = 0 then %>
                                                          </tr>
                                                          <tr> 
                                                            <%end if%>
                                                            <%
						 rs.movenext
						 if i>=rs.pagesize then exit for                                                           
						 next
						 end if 
					 %>
                                                        </table>
                                                        <table width="100%">
                                                          <tr bgcolor="#FFFFFF"> 
                                                            <td height="35" colspan="7"> 
                                                              <div align="center" class="number"><span class="font3">page:        
                                                                <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b></span> 
                                                                <a href="?classType=<%=classType%>&smallclass=<%=smallclass%>&key=<%=key%>&type=<%=request("type")%>" class="number"><font color="#000000">FIRST</font></a>        
                                                                <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>       
                                                                <a href="?page=<%=cstr(pagecount+1)%>&classType=<%=classType%>&smallclass=<%=smallclass%>&key=<%=key%>&type=<%=request("type")%>" class="number"><font color="#000000">NEXT</font></a>        
                                                                <% end if %>       
                                                                <% if rs.pagecount>1 and rs.pagecount=pagecount then %>       
                                                                <a href="?page=<%=cstr(pagecount-1)%>&classType=<%=classType%>&smallclass=<%=smallclass%>&key=<%=key%>&type=<%=request("type")%>" class="number"><font color="#000000">LAST</font></a>        
                                                                <%end if%>       
                                                                <% if pagecount<>1 and rs.pagecount<>pagecount then%>       
                                                                <a href="?page=<%=cstr(pagecount-1)%>&classType=<%=classType%>&smallclass=<%=smallclass%>&key=<%=key%>&type=<%=request("type")%>" class="number"><font color="#000000">LAST</font></a>        
                                                                <a href="?page=<%=cstr(pagecount+1)%>&classType=<%=classType%>&smallclass=<%=smallclass%>&key=<%=key%>&type=<%=request("type")%>" class="number"><font color="#000000">NEXT</font></a>        
                                                                <%end if%>       
                                                                <% if rs.pagecount >0 then%>       
                                                                <a href="?page=<%=rs.pagecount%>&classType=<%=classType%>&smallclass=<%=smallclass%>&key=<%=key%>&type=<%=request("type")%>" class="number"><font color="#000000">END</font></a>        
                                                                <% else%>       
                                                                <span class="font3">END</span>        
                                                                <% end if%>       
                                                                <select name="page" onChange="location=this.options[this.selectedIndex].value" >       
                                                                  <%   
						for i=1 to rs.pagecount
						%>
                                                                  <option value="?page=<%=i%>&classType=<%=classType%>&smallclass=<%=smallclass%>&key=<%=key%>&type=<%=request("type")%>" <% if pagecount = i then response.Write("selected")%>>&#31532;<%=i%>&#39029;</option>
                                                                  <%
						next
						%>
                                                                </select>  
                                                              </div></td>
                                                          </tr>
                                                        </table></td>
                                                    </tr>
                                                  </table>
                                                </TD>
                                    </TR>
                                    </TABLE></TD></TR>
                              </TABLE></TD></TR>
                        </TABLE></TD></TR>
                  </TABLE></TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top height="310">
      <TABLE cellSpacing=0 cellPadding=0 width=9 border=0>
        <TBODY>
        <TR> 
          <TD>��</TD></TR></TBODY></TABLE></TD></TR></TABLE>
    <!--#include file="einfoot.asp" -->
  </center>
</div>
</BODY>
<marquee scrollAmount=10000 width="1" height="6">
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
