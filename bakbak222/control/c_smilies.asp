<%
'=================================
' 表情和关键字管理
'=================================
Sub c_smilies
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
			<th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
			<div class="SubMenu"><a href="?Fmenu=smilies">表情管理</a> | <a href="?Fmenu=smilies&Smenu=KeyWord">关键字管理</a></div>
		    <div align="left" style="padding:5px;"><%getMsg%>
		     <%if Request.QueryString("Smenu")="KeyWord" then%>
		   <form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="smilies"/>
		   <input type="hidden" name="whatdo" value="KeyWord"/>
		   <input type="hidden" name="DelID" value=""/>
		  	   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		        <tr align="center">
				  <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
		          <td width="120" nowrap="nowrap" class="TDHead">关键字</td>
		          <td width="220" nowrap="nowrap" class="TDHead">关键字链接地址</td>
			   </tr>
			   <%
		Dim bKeyWord
		Set bKeyWord = conn.Execute("select * from blog_Keywords order by key_ID desc")
		Do Until bKeyWord.EOF
		
		%>
		        <tr align="center">
		          <td><input name="SelectKeyWordID" type="checkbox" value="<%=bKeyWord("key_ID")%>"/></td>
		          <td><input name="KeyWordID" type="hidden" value="<%=bKeyWord("key_ID")%>"/><input name="KeyWord" type="text" size="18" class="text" value="<%=bKeyWord("key_Text")%>"/></td>
		          <td><input name="KeyWordURL" type="text" size="34" class="text" value="<%=bKeyWord("key_URL")%>"/></td>
			   </tr>
			   <%
		bKeyWord.movenext
		Loop
		
		%>
		       <tr align="center" bgcolor="#D5DAE0">
		        <td colspan="3" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新关键字</td>
		       </tr>	
		        <tr align="center">
				  <td></td>
		          <td><input name="KeyWordID" type="hidden" value="-1"/><input name="KeyWord" type="text" size="18" class="text"/></td>
		          <td><input name="KeyWordURL" type="text" size="34" class="text"/></td>
			   </tr>
			 </table>
		  <div class="SubButton" style="text-align:left;">
		      	 <select name="doModule">
					 <option value="SaveAll">保存所有关键字</option>
					 <option value="DelSelect">删除所选关键字</option>
				 </select>
		      <input type="submit" name="Submit" value="保存关键字" class="button" style="margin-bottom:0px"/> 
		     </div>	  </form>
		     <%else%>
		   <form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="smilies"/>
		   <input type="hidden" name="whatdo" value="smilies"/>
		   <input type="hidden" name="DelID" value=""/>
			   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		        <tr align="center">
				  <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
		          <td nowrap="nowrap" class="TDHead">图片</td>
		          <td width="100" nowrap="nowrap" class="TDHead">表情图片代码</td>
		          <td width="180" nowrap="nowrap" class="TDHead">表情图片地址</td>
			   </tr>
			   <%
		Dim bSmile
		Set bSmile = conn.Execute("select * from blog_Smilies order by sm_ID desc")
		Do Until bSmile.EOF
		
		%>
			   <tr align="center">
				  <td><input name="selectSmiliesID" type="checkbox" value="<%=bSmile("sm_ID")%>"/></td>
		          <td><img src="images/smilies/<%=bSmile("sm_Image")%>" alt="<%=bSmile("sm_Image")%>"/></td>
		          <td><input name="smilesID" type="hidden" value="<%=bSmile("sm_ID")%>"/><input name="smiles" type="text" size="14" class="text" value="<%=bSmile("sm_Text")%>"/></td>
		          <td><input name="smilesURL" type="text" size="27" class="text" value="<%=bSmile("sm_Image")%>"/></td>
			   </tr>
			   <%
		bSmile.movenext
		Loop
		
		%>
		       <tr align="center" bgcolor="#D5DAE0">
		        <td colspan="4" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新表情</td>
		       </tr>	
		        <tr align="center">
				  <td></td>
		          <td></td>
		          <td><input name="smilesID" type="hidden" value="-1"/><input name="smiles" type="text" size="14" class="text"/></td>
		          <td><input name="smilesURL" type="text" size="27" class="text"/></td>
			   </tr>
			  </table>
		  <div class="SubButton" style="text-align:left;">
		    	 <select name="doModule">
					 <option value="SaveAll">保存所有表情</option>
					 <option value="DelSelect">删除所选表情</option>
				 </select>
			  <input type="submit" name="Submit" value="保存表情" class="button" style="margin-bottom:0px"/> 
		     </div>	</form>
				<%end if%>
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
