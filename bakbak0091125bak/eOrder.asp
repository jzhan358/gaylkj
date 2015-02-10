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
          <TD>¡¡</TD></TR></TBODY></TABLE></TD>
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
                      <TD height=20><!-- //×óÀ¸²Ëµ¥ -->
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
rs.OPEN sql,Conn,1,1
if rs.eof and rs.bof then
Response.Write"no product type"
else
do while not rs.eof

sql = "SELECT * FROM P_class_small where p_type ='"&rs("p_type")&"'"
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.OPEN sql,Conn,1,1
if rs2.recordcount > 0 then 
Response.Write"&nbsp;<A border=0 style='CURSOR: hand' id=Img"&rs("p_id")&" onclick=ToggleDisplay(Img"&rs("p_id")&",Tree"&rs("p_id")&")><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type_e")&"</b></a>"
else
Response.Write"&nbsp;<A href='eProduct.asp?classtype="&rs("p_type_e")&"' class='six'><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type_e")&"</b></a>"
end if
counts=conn.execute("Select count(*) from P_info where p_other = 0 and p_type='"&rs("p_type")&"'")(0)
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
                                <td width="100%">&nbsp; ¡ó&nbsp;<a class="number" href='eProduct.asp?classtype=<%=rs2("p_type_e")%>&smallclass=<%=rs2("p_small_type_e")%>'><%=rs2("p_small_type_e")%></a></td>          
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
                                        <TD height=3><!--<img border="0" src="images/2Left_About2.gif">--></TD>
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
                                                   <!-- <DIV align=center>
                                                      <table width="100%" height="34" border="0" cellpadding="0" cellspacing="0">
                                                        <form name="form1" method="post" action="Demo_eProduct.asp">
                                                            <td width="124" align="left"><input name="key" type="text" class="bd5" id="key" size="15"></td>
                                                            <td width="48" align="left"><input name="imageField" type="image" src="images/esubmit.gif" border="0"></td>
                                                        </form>
                                                      </table>
                                                    </DIV>--></TD>
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
                                <TD>¡¡</TD></TR></TBODY></TABLE></TD>
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
                                color=#ffffff><a href="eindex.asp"><FONT 
                                color=#ffffff>Home</FONT></a>&nbsp; |&nbsp; <a href="eintro.asp"><FONT                                          
                                color=#ffffff>Company</FONT></a>&nbsp; |&nbsp; <a href="eNews.asp"><FONT         
                                color=#ffffff>News</FONT></a>&nbsp;                                             
                                |&nbsp; <a href="eProduct.asp"><FONT                                           
                                color=#ffffff>Products</FONT></a>&nbsp; |&nbsp;<a href="eGuest.asp"><FONT                                          
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
                                <TD width=43><img border="0" src="index/order.gif"></TD>
                                <TD width="100%">
                                <P align=right>¡¡</P></TD></TR></TBODY></TABLE></TD>
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

	id=request.QueryString("id")
	title=""
	if isnumeric(id) then
	id=cint(id)
	if id>0 and id<>"" then
		set rs=conn.execute("select p_name_e from p_info where p_id="&id)
		if not rs.eof then
		title=getInnerText(rs("p_name_e"))	
		else
		id=""	
		end if
		rs.close
		set rs=nothing
	else
	id=""
	end if
	else
	id=""
	end if

%>

                                                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                                    <tr>
                                  <TD height=5>
                                    <table border="0" width="100%" cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td width="100%" height="8"></td>
                                      </tr>
                                      <tr>
                                        <td width="100%"><TABLE width="100%" height="327" border=0 align="center" cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height=5></TD>
    </TR>
    <TR> 
      <TD height=278 align="center" valign="middle">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
          <form action="Demo_eOrderPost.asp" method="post" name="form0" target="_self">
            <input name="bookpost" type="hidden" value="bookpost">
            <tr> 
              <td width="40" align="left" rowspan="11"></td>
              <td width="146" height="25" align="left">&nbsp;&nbsp;Name:<span class="style6">*</span></td>
              <td width="429" height="25" align="left"><input name="company" type="text" id="company2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="left" valign="top">&nbsp;&nbsp;Address:<span class="style6">*</span></td>
              <td height="25" align="left"><input name="addr" type="text" id="addr2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="left">&nbsp;&nbsp;Telphone:<span class="style6">*</span></td>
              <td height="25" align="left"><input name="tel" type="text" id="tel2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="left">&nbsp;&nbsp;Fax:</td>
              <td height="25" align="left"><input name="fax" type="text" id="fax2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="left">&nbsp;&nbsp;Linkman:</td>
              <td height="25" align="left"><input name="linkman" type="text" id="linkman2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="left">&nbsp;&nbsp;Product:<span class="style6">*</span></td>
              <td height="25" align="left"><input name="proname" type="text" id="proname2" value="<%=title%>" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="left">&nbsp;&nbsp;Spec:<span class="style6">*</span></td>
              <td height="25" align="left"><input name="proxh" type="text" id="proxh2" value="<%=id%>" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="12" align="left">&nbsp;&nbsp;Number:<span class="style6">*</span></td>
              <td height="-2" align="left"><input name="num" type="text" id="num2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="12" align="left" valign="top">&nbsp;&nbsp;Content:</td>
              <td height="0" align="left"><textarea name="bz" cols="45" rows="8" id="textarea2" style="font-size:12px"></textarea></td>
            </tr>
            <tr> 
              <td height="25" align="left">&nbsp;&nbsp;Date:</td>
              <td height="25" align="left"><input name="ordertime" type="text" id="ordertime2" value="<%=date()%>" size="40" maxlength="50"></td>
            </tr>
            <tr align="center" valign="middle"> 
              <td height="30" colspan="2"><input type="submit" name="Submit2" value=" OK " class="button">&nbsp;&nbsp;&nbsp;    
                <input type="reset" name="Submit2" value="cancel" class="button"></td>
            </tr>
          </form>
        </table>
      </TD>
    </TR>
    <TR> 
      <TD height="0" background=images/main_heng.gif>&nbsp;&nbsp;&nbsp;&nbsp;    
        Note: bringing &quot;*&quot; number is will fill option!</TD>    
    </TR>
    <TR> 
      <TD 
              height=13 align=right>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE></td>
                                      </tr>
                                      <tr>
                                        <td width="100%"></td>
                                      </tr>
                                    </table>
                                  </TD>
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
          <TD>¡¡</TD></TR></TBODY></TABLE></TD></TR></TABLE>
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
