<%
'=================================
' 日志分类管理
'=================================
Sub c_article
%>
<%getMsg%>
    <table width="100%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" class="CContent">
	    <tr>
		    <td bgcolor="#FFFFFF" class="CTitle">日志管理v0.2 For Pjblog3 </td>
   		</tr>
	    <%IF Request.QueryString("type")="LogMG" Then%>
	    <tr>
		    <td align="center" bgcolor="#FFFFFF" height="48">
		    <%
		        If Request.form("moveto")=1 Then
		            Dim Log_Dele,Log_source_ID
		            Log_Dele=split(Request.form("Log_Dele"),", ")
		            for i=0 to ubound(Log_Dele)
		                Log_source_ID=conn.execute("select log_CateID from blog_Content where log_ID="&Log_Dele(i))(0)
		                conn.execute ("update blog_Content set log_CateID="&Request.form("source")&" where log_ID="&Log_Dele(i))
		                conn.execute ("update blog_Category set cate_count=cate_count+1 where cate_ID="&Request.form("source"))
		                conn.execute ("update blog_Category set cate_count=cate_count-1 where cate_ID="&Log_source_ID)
		            next
		             session(CookieName&"_ShowMsg") = True
		            session(CookieName&"_MsgText") = "日志移动成功！"
			        FreeMemory
		            RedirectUrl("ConContent.asp?Fmenu=Article")
		        Else
		            Log_Dele=split(Request.form("Log_Dele"),", ")
		            dim fso
		            Set fso = CreateObject("Scripting.FileSystemObject")
		            for i=0 to ubound(Log_Dele)
		                Log_source_ID=conn.execute("select log_CateID from blog_Content where log_ID="&Log_Dele(i))(0)
		                conn.execute ("update blog_Category set cate_count=cate_count-1 where cate_ID="&Log_source_ID)
		                conn.execute("DELETE * from blog_Content where log_ID="&Log_Dele(i))
		                if fso.FileExists(server.MapPath("/article/"& Log_Dele(i) &".htm")) then
		                    fso.DeleteFile(server.MapPath("/article/"& Log_Dele(i) &".htm"))
		                end if
		            next
		             session(CookieName&"_ShowMsg") = True
		            session(CookieName&"_MsgText") = "日志删除成功！"
			        FreeMemory
					Session(CookieName&"_LastDo")="DelArticle"
		            RedirectUrl("ConContent.asp?Fmenu=Article&cate_ID="&Request.form("cate_ID")&"")
		       End If
		    %>
		    </td>
	    </tr>
    <%Else%>
	    
	    <tr>
		    <td class="CPanel">
		    <div class="SubMenu">
		    <%
		        Dim Log_cate,Log_cateid
				Log_cateid=Request("cate_ID")
				If Log_cateid="" then
					Log_cateid=0
				End if
		        Set Log_cate=Server.CreateObject("ADODB.RecordSet")
		        Sql="select * from blog_Category where not cate_OutLink"
		        Log_cate.Open Sql,conn,1,1
		        If Log_cate.eof and Log_cate.bof then
		            response.write "暂未添加分类！"
		        Else
					Dim Log_c
					Set Log_c=Server.CreateObject("ADODB.RecordSet")
					Log_c.Open "SELECT COUNT(*) as Mycount FROM blog_Content",Conn
		                If Log_cateid>0 then
							Response.write "<a href=ConContent.asp?Fmenu=Article&Smenu=>查看全部("&Log_c("MyCount")&")</a> | "
		                Else
							Response.write "<a href=ConContent.asp?Fmenu=Article&Smenu=><b>查看全部("&Log_c("MyCount")&")</b></a> | "
		                End If
					Log_c.Close
					Set Log_c=nothing
		            Do While Not Log_cate.eof
		                If Log_cate("cate_ID")=int(Log_cateid) then
		                    response.write "<a href='ConContent.asp?Fmenu=Article&cate_ID="&Log_cate("cate_ID")&"&Smenu='><b>"&Log_cate("cate_Name")&"</b></a>("&Log_cate("cate_Count")&") | "
		                Else
		                    response.write "<a href='ConContent.asp?Fmenu=Article&cate_ID="&Log_cate("cate_ID")&"&Smenu='>"&Log_cate("cate_Name")&"</a>("&Log_cate("cate_Count")&") | "
		                End If
		            Log_cate.MoveNext
		            Loop	
		        End If
		        Log_cate.Close
		        Set Log_cate=Nothing
		    %></div>
		    
		    </td>
		  </tr>
		    <tr>
			    <td align="center" valign="top" class="CPanel" style="padding-left:20px;padding-right:20px">
			    <a href="blogpost.asp" target="_blank" style="float:left;margin-bottom:2px;"><b><img style="margin: 0px 2px -3px 0px" src="images/add.gif" border="0" />发表新日志</b></a>
				<div style="float:right;margin-right:6px;">
					排序: <b>时间</b> <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=1"><img src="images/sort_up.gif" border=0/></a>
					 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=2"><img src="images/sort_down.gif" border=0/></a>
					  <span style="color:#999">|</span> <b>访问</b> <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=3"><img src="images/sort_up.gif" border=0/></a>
					<a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=4"><img src="images/sort_down.gif" border=0/></a>
					 <span style="color:#999">|</span> <b>评论</b> <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=5"><img src="images/sort_up.gif" border=0/></a>
					  <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=6"><img src="images/sort_down.gif" border=0/></a>
					   <span style="color:#999">|</span> <b>引用</b> <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=7"><img src="images/sort_up.gif" border=0/></a> 
					   <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=8"><img src="images/sort_down.gif" border=0/></a> 
				</div>
			    
					<form action="ConContent.asp?Fmenu=Article&type=LogMG" method="post" name="ph_Category" id="ph_Category" style="margin:0px;">
			               <input type="hidden" name="doModule" value="DelSelect"/>
			               <input type="hidden" name="cate_ID" value="<%=Request.QueryString("cate_ID")%>"/>
				    <%
				        If CheckStr(Request.QueryString("Page"))<>Empty Then
				            Curpage=CheckStr(Request.QueryString("Page"))
				            If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
				        Else
				            Curpage=1
				        End If
				    
				    Dim Log_List
				    Set Log_List=Server.CreateObject("ADODB.RecordSet")
				    
				    If Request.QueryString("cate_ID")<>Empty Then
				        Sql="select log_ID,log_CateID,log_Title,log_PostTime,log_CommNums,log_QuoteNums,log_ViewNums,cate_ID,cate_Name from blog_Content c inner join blog_Category l on c.log_CateID=l.cate_ID Where log_CateID="&Request.QueryString("cate_ID")&""
				    Else
				        Sql="select log_ID,log_CateID,log_Title,log_PostTime,log_CommNums,log_QuoteNums,log_ViewNums,cate_ID,cate_Name from blog_Content c inner join blog_Category l on c.log_CateID=l.cate_ID"
				    End If
				    If Request.QueryString("Log_sort")<>Empty Then
				        Select Case Request.QueryString("Log_sort")
				            Case 1
				                Sql=Sql&" order by log_PostTime"
				            Case 2
				                Sql=Sql&" order by log_PostTime desc"
				            Case 3
				                Sql=Sql&" order by log_ViewNums"
				            Case 4
				                Sql=Sql&" order by log_ViewNums desc"
				            Case 5
				                Sql=Sql&" order by log_CommNums"
				            Case 6
				                Sql=Sql&" order by log_CommNums desc"
				            Case 7
				                Sql=Sql&" order by log_QuoteNums"
				            Case 8
				                Sql=Sql&" order by log_QuoteNums desc"
				        End Select
				    Else
				        Sql=Sql&" order by log_ID desc"	
				    End If
				    
				    Log_List.Open Sql,conn,1,1
				    
				    If not Log_List.eof Then
				        Dim Log_PageCM
				        Log_PageCM=0
				        Log_List.PageSize=15
				        Log_List.AbsolutePage=CurPage
				        Dim Log_List_nums
				        Log_List_nums=Log_List.RecordCount
						Dim urlLink    '根据动静态判断输出连接  JieLiao
				    %>
				    <table border="0" cellpadding="2" cellspacing="1" class="TablePanel" width="100%">
				    <tr>
					    <td class="TDHead" nowrap>选择</td>
					    <td  class="TDHead" width="100%">标题</td>
					    <td  class="TDHead" nowrap width="150">发布时间</td>
					    <td  class="TDHead" nowrap width="50">评论</td>
					    <td  class="TDHead" nowrap width="50">引用</td>
					    <td  class="TDHead" nowrap width="50">查看</td>
					    <td  class="TDHead" nowrap width="50">操作</td>
				    </tr>
				    <%
					Do Until Log_List.EOF OR Log_PageCM=Log_List.PageSize
						if blog_postFile = 2 then
							urlLink = "article/"&Log_List(0)&".htm"
						else 
							urlLink = "article.asp?id="&Log_List(0)
						end if

				    %>
				    <tr bgcolor="#FFFFFF">
				    <td align="center"><input name="Log_Dele" type="checkbox" id="Log_Dele" value=<%=Log_List(0)%>></td>
				    <td>
				    <%
				    If Request.QueryString("cate_ID")=Empty Then
				        response.write "<a href=ConContent.asp?Fmenu=Article&cate_ID="&Log_List(1)&">【"&Log_List(8)&"】</a>"
				    End If
				    %>
				    <a target="_blank" href="<%=urlLink%>"><%=Log_List(2)%></a></td>
				    <td><%=Log_List(3)%></td>
				    <td align="center">
				    <%
				    If Log_List(4)>0 then
				    %>
				    <a href="<%=urlLink%>#comm_top" target="_blank"><%=Log_List(4)%></a>
				    <%
				    Else
				    %>
				    0
				    <%End If
				    %>
				    </td>
				    <td align="center"><%=Log_List(5)%></td>
				    <td align="center"><%=Log_List(6)%></td>
				    <td align="center"><a target="_blank" href="blogedit.asp?id=<%=Log_List(0)%>"><img style="margin: 0px 2px -3px 0px" src="images/icon_edit.gif" border="0" />编辑</a>
				    </select>
				    </td>
				    </tr>
				    <%
				    Log_List.MoveNext
				    Log_PageCM=Log_PageCM+1
				    Loop
				    %>
				    <tr><td colspan="7" bgcolor="#ffffff">
			            <input type="button" value="全选" onClick="checkAll()" class="button" style="margin:0px;margin-bottom:5px;margin-right:6px"/>
			            <input type="button" value="删除所选内容" onClick="DelComment()" class="button" style="margin:0px;margin-bottom:5px;"/>
			            <input type="hidden" value="0" name="moveto">
			           <input type="submit" value="将所选内容移至" onClick="moveto.value=1" class="button" style="margin:0px;margin-bottom:5px;"/>
			           <select name="source"  style="margin:0px;margin-bottom:5px;">
					    <%
					    Dim Log_CategoryListDB,Log_CateInOpstions
					            set Log_CategoryListDB=conn.execute("select * from blog_Category order by cate_local asc, cate_Order desc")
					             do while not Log_CategoryListDB.eof
					              if not Log_CategoryListDB("cate_OutLink") then
					               Log_CateInOpstions=Log_CateInOpstions&"<option value="""&Log_CategoryListDB("cate_ID")&""">&nbsp;&nbsp;"&Log_CategoryListDB("cate_Name")&" ["&Log_CategoryListDB("cate_count")&"]</option>"
					              end if
					              Log_CategoryListDB.movenext
					             loop
					             set Log_CategoryListDB=nothing
					    %>
					        <%=Log_CateInOpstions%>
					       </select>
				    </td></tr>				   
			   	 	</table>

    	    <%
	        response.write "<div class=""pageContent"">"&MultiPage(Log_List_nums,Log_List.PageSize,CurPage,"?Fmenu=Article&Log_sort="&Request.QueryString("Log_sort")&"&cate_ID="&Request.QueryString("cate_ID")&"&","","float:left")&"</div>"
	    Else
	        response.write ("该分类暂无日志不存在！")
	    End If
	    Log_List.close
	    Set Log_List=Nothing
	    %>
    <%End IF%>
    </form>
   		 </td>
   	 </tr>
    </table>
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
