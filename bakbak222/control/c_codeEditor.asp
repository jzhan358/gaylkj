<%
'=================================
' 代码编辑器
'=================================
Sub c_codeEditor
	dim source,filePath,urlPath,fileInfo,Arrfiles,AttPath,Arrfile,fileItem,fileItemInfo,fileExt,referer,fileType,loadFile
    Set fileExt = Server.CreateObject("Scripting.Dictionary")
    fileExt.Add ".css" , "css"
    fileExt.Add ".html" , "html" 
    fileExt.Add ".htm" , "html"
    fileExt.Add ".asp" , "html"
    fileExt.Add ".xml" , "html"
    
	filePath = Request.QueryString("file")
	referer = Request.QueryString("referer")

	fileInfo = getFileInfo(filePath)
	fileType = fileExt(fileInfo(1))
	
	
	urlPath = siteURL & replace(filePath,"\","/")
	
	loadFile = LoadFromFile(filePath)
	
	source = loadFile(1)
    source = Replace(source, ">", "&gt;")
    source = Replace(source, "<", "&lt;")
    
	AttPath = left(filePath,InStrRev(filePath,"\"))
	Arrfiles = Split(getPathList(AttPath)(1), "*")
	%>
	<script type="text/javascript">
		var goBack = function(type) {
			switch(type){
				case "skins" :
					location = 'ConContent.asp?Fmenu=Skins&Smenu=';
					break;
				case "plugins" :
					location = 'ConContent.asp?Fmenu=Skins&Smenu=PluginsInstall';
					break;
				default:
					location = 'ConContent.asp?Fmenu=welcome'
			}
		}
		
		var saveFile = function(){
			document.getElementById("cCode").value = myCode.getCode();
			document.getElementById("cPoster").submit();
		}
	</script>
	<%getMsg%>
	<form id="cPoster" name="codePoster" action="ConContent.asp" method="post" style="margin:0px">
			<input type="hidden" name="action" value="codeEditor"/>
			<input type="hidden" name="whatdo" value="save"/>
			<input type="hidden" name="referer" value="<%=referer%>"/>
			<input type="hidden" name="filePath" value="<%=filePath%>"/>
			<input type="hidden" name="code" id="cCode" value=""/>
	</form>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		<tr>
			<th class="CTitle"><%=categoryTitle%></th>
		</tr>
		<tr>
			<td class="CPanel">
				<div class="SubMenu2">
					<input type="button" name="Submit" value="返回" onclick="goBack('<%=referer%>')" class="button" style="float:right"/>
					<input type="button" name="Submit" value="保存文件" onclick="saveFile()" class="button" style="float:right"/>
					<%=getFileIcons(fileInfo(1))%><a href="<%=urlPath%>"><%=urlPath%></a>
					<div style="padding-left:18px;"><span style="color:#999">文件大小:</span> <%=fileInfo(0)%> &nbsp;&nbsp;<span style="color:#999">文件修改时间:</span> <%=fileInfo(4)%> &nbsp;&nbsp;<span style="color:#999">文件类型:</span> <%=fileInfo(3)%></div>
				</div>
				<div style="float:left;border:1px solid #808080;width:160px;height:385px;margin-right:4px;">
					<div class="SubMenu" style="position:absolute;z-index:1px;width:140px;margin:0;padding-top:0;padding-left:4px;cursor:default;text-overflow:ellipsis " title="<%=AttPath%>"><img border="0" src="images/file/folder.gif" style="margin:4px 3px -3px 0px"/><%=AttPath%></div>
					<div style="height:353px;*height:383px;overflow-y:scroll;overflow-x:hidden;padding:2px;padding-top:30px;">
<%
		'列举出文件
		For Each Arrfile in Arrfiles
			fileItem = AttPath&Arrfile
			fileItemInfo = getFileInfo(fileItem)
			if Not isEmpty(fileExt(fileItemInfo(1))) then
				response.write "<a href=""?Fmenu=CodeEditor&file="&fileItem&"&referer="&referer&""">" &getFileIcons(fileItemInfo(1)) & fileItemInfo(7) & "</a><br/>"
			end if
		  '  response.Write "<a href="""&AttPath&"/"&Arrfile&""" target=""_blank"">"&getFileIcons(getFileInfo(AttPath&"/"&Arrfile)(1))&Arrfile&"</a>&nbsp;&nbsp;"&getFileInfo(AttPath&"/"&Arrfile)(0)&" | "&getFileInfo(AttPath&"/"&Arrfile)(2)&" | "&getFileInfo(AttPath&"/"&Arrfile)(3)&"<br>"
		Next
%>
					</div>
				</div>
				<div style="width:790px;height:380px;*height:385px;border:1px solid #808080;padding:4px 4px 0 4px;float:left">
					<div class="SubMenu" style="padding-top:0;padding-left:8px;margin:-4px -4px 4px -4px;">
						<a href="javascript:;" onclick="myCode.toggleEditor()" title="开启/关闭 代码高亮"><img border="0" src="images/code.gif" style="margin:4px 3px -3px 0px"/>代码高亮</a> | 
						<a href="javascript:;" onclick="myCode.toggleAutoComplete()" title="开启/关闭 自动完成"><img border="0" src="images/auto.gif" style="margin:4px 3px -3px 0px"/>自动完成</a> | 
						<a href="javascript:;" onclick="myCode.toggleLineNumbers()" title="显示/隐藏 行号"><img border="0" src="images/line.gif" style="margin:4px 3px -3px 0px"/>行号</a>
					</div>
			  	  	<textarea id="myCode" class="codepress <%=fileType%> linenumbers-on" wrap="on" style="width:100%;height:346px;"><%=source%></textarea>
				</div>

			</td>
		</tr>
	</table>
	<script src="common/codepress/codepress.js" type="text/javascript"></script> 
	<%
	set fileExt = nothing
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
