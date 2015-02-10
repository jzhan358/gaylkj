<!--#include file="config.asp"-->
<!--#include file="admin/inc/format.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>demo - product </title>
</head>

<body>
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td width="19%" valign="top"> <TABLE borderColor=#111111 cellSpacing=0 cellPadding=2 width="100%" border=0 class="font22">
        <tr> 
          <TD width=100% height="16"> <script language="JavaScript" type="text/JavaScript">
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

sql = "SELECT * FROM P_class_small where p_type ='"&rs("p_type")&"'"
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.OPEN sql,Conn,1,1
if rs2.recordcount > 0 then 
Response.Write"&nbsp;<A border=0 style='CURSOR: hand' id=Img"&rs("p_id")&" onclick=ToggleDisplay(Img"&rs("p_id")&",Tree"&rs("p_id")&")><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type_e")&"</b></a>"
else
Response.Write"&nbsp;<A href=eProduct.asp?classtype="&rs("p_type_e")&" class='six'><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type_e")&"</b></a>"
end if
counts=conn.execute("Select count(*) from P_info where p_other = 1 and p_type='"&rs("p_type")&"'")(0)
Response.Write"(<FONT COLOR='800000'>"&counts&"</FONT>)<br>"
%>
          </td>
        </tr>
        <tr> 
          <td><img src="images/sortxiuxian.jpg"></td>
        </tr>
        <tr>
          <td> <div id=Tree<%=rs("p_id")%>> 
              <%
do while not rs2.eof
%>
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="100%">&nbsp;&nbsp;&nbsp;<a class="number" href='eProduct.asp?classtype=<%=rs2("p_type_e")%>&smallclass=<%=rs2("p_small_type_e")%>'><%=rs2("p_small_type_e")%></a></td>
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
            <SCRIPT language=javascript>FoldAll();</SCRIPT> </td>
        </tr>
      </table></td>
    <td width="81%" valign="top"> 
      <table width="570" border="0" cellpadding="0" cellspacing="0">
        <tr align="left"> 
          <td width="75">search:</td>
          <td class="font7"> <%=request("classtype")%></td>
          <td width="268" align="left" valign="bottom"> <table width="268" height="34" border="0" cellpadding="0" cellspacing="0">
              <form name="form1" method="post" action="eProduct.asp">
                <tr> 
                  <td width="96" align="right" class="font18">search：</td>
                  <td width="124" align="left"><input name="key" type="text" class="bd5" id="key"></td>
                  <td width="48" align="left"><input name="imageField" type="image" src="images/submit.gif" border="0"></td>
                </tr>
              </form>
            </table></td>
          <td width="19" align="right">&nbsp; </td>
        </tr>
      </table>

<%
ProID = strreplace(clng(request("ProID")))
sql = "select * from P_info where  p_id = "& ProID
set rs = conn.execute(sql)
if not(rs.eof and rs.bof) then
%>	
      <table width="570" height="239" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="801">Product:<%=rs("p_name_e")%></td>
        </tr>
        <tr> 
          <td>model:<%=rs("p_spec_e")%></td>
        </tr>
        <tr>
          <td>product type:<%=rs("p_type_e")%> - <%=rs("p_small_type_e")%></td>
        </tr>
        <tr>
          <td>content:<%=rs("P_jianjie_e")%></td>
        </tr>
        <tr>
          <td>more content:<%=rs("p_epitome_e")%></td>
        </tr>
      </table>
	  <%
end if
%>      

	  </td>
  </tr>
</table>
</body>
</html>
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
