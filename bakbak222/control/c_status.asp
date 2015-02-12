<%
'=================================
' 服务器信息
'=================================
Sub c_status
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
			<th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
		    <div align="left" style="padding:5px;line-height:150%">
		     <b>软件版本:</b> PJBlog3 v<%=blog_version%> - <%=DateToStr(blog_UpdateDate,"mdy")%><br/>
		     <b>服务器时间:</b> <%=DateToStr(Now(),"Y-m-d H:I A")%><br/>
		     <b>服务器物理路径:</b> <%=Request.ServerVariables("APPL_PHYSICAL_PATH")%><br/>
		     <b>服务器空间占用:</b> <%=GetTotalSize(Request.ServerVariables("APPL_PHYSICAL_PATH"),"Folder")%><br/>
		     <b>服务器CPU数量:</b> <%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%><br/>
		     <b>服务器IIS版本:</b> <%=Request.ServerVariables("SERVER_SOFTWARE")%><br/>
		     <b>脚本超时设置:</b> <%=Server.ScriptTimeout%><br/>
		     <b>脚本解释引擎:</b> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %><br/>
		     <b>服务器操作系统:</b> <%=Request.ServerVariables("OS")%><br/>
		     <b>服务器IP地址:</b> <%=Request.ServerVariables("LOCAL_ADDR")%><br/>
		     <b>客户端IP地址:</b> <%=Request.ServerVariables("REMOTE_ADDR")%><br/><br/>
		     
		     <b>关键组件:</b> (缺少关键组件的服务器会对PJBlog3运行有一定影响)<br/>
		     <b>　－ Scripting.FileSystemObject 组件:</b> <%=DisI(CheckObjInstalled("Scripting.FileSystemObject"))%><br/>
		     <b>　－ MSXML2.ServerXMLHTTP 组件:</b> <%=DisI(CheckObjInstalled("MSXML2.ServerXMLHTTP"))%><br/>
		     <b>　－ Microsoft.XMLDOM 组件:</b> <%=DisI(CheckObjInstalled("Microsoft.XMLDOM"))%><br/>
		     <b>　－ ADODB.Stream 组件:</b> <%=DisI(CheckObjInstalled("ADODB.Stream"))%><br/>
		     <b>　－ Scripting.Dictionary 组件:</b> <%=DisI(CheckObjInstalled("Scripting.Dictionary"))%><br/>
		     <br/>
		     
		     <b>其他组件: </b>(以下组件不影响PJBlog3运行)<br/>
		     <b>　－ Msxml2.ServerXMLHTTP.5.0 组件:</b> <%=DisI(CheckObjInstalled("Msxml2.ServerXMLHTTP.5.0"))%><br/>
		     <b>　－ Msxml2.DOMDocument.5.0 组件:</b> <%=DisI(CheckObjInstalled("Msxml2.DOMDocument.5.0"))%><br/>
		     <b>　－ FileUp.upload 组件:</b> <%=DisI(CheckObjInstalled("FileUp.upload"))%><br/>
		     <b>　－ JMail.SMTPMail 组件:</b> <%=DisI(CheckObjInstalled("JMail.SMTPMail"))%><br/>
		     <b>　－ GflAx190.GflAx 组件:</b> <%=DisI(CheckObjInstalled("GflAx190.GflAx"))%><br/>
		     <b>　－ easymail.Mailsend 组件:</b> <%=DisI(CheckObjInstalled("easymail.Mailsend"))%><br/>
		    </div>
		</td></tr></table>
<%
end Sub
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
