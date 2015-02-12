﻿<!--#include file="BlogCommon.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="plugins.asp" -->
<!--内容-->
<%
'==================================
'  友情连接
'    更新时间: 2005-12-11
'==================================
%>
 <!--内容-->
  <div id="Tbody">
   <div id="mainContent">
     <div id="innermainContent">
       <div id="mainContent-topimg"></div>
       	 <div id="Content_ContentList" class="content-width">
       	 
         <%
If request.Form("action") = "postLink" Then
    Dim link_Name, link_URL, link_Image, linkCount, linkDB, linkvalidate
    link_Name = checkURL(checkstr(request.Form("link_Name")))
    link_URL = checkURL(checkstr(request.Form("link_URL")))
    link_Image = checkURL(checkstr(request.Form("link_Image")))
    linkvalidate = checkURL(checkstr(request.Form("link_validate")))
    If CStr(LCase(Session("GetCode")))<>CStr(LCase(linkvalidate)) And Not stat_Admin Then
        showmsg "友情链接发表出错", "<b>验证码有误，请返回重新输入</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a>", "ErrorIcon", ""
    End If
    If Len(link_Name)<1 Then
        showmsg "友情链接发表出错", "<b>网站名称不能为空！</b><br/><a href=""javascript:history.go(-1);"">返回</a>", "ErrorIcon", ""
    End If
    If Len(link_URL)<1 Then
        showmsg "友情链接发表出错", "<b>网站地址不能为空！</b><br/><a href=""javascript:history.go(-1);"">返回</a>", "ErrorIcon", ""
    End If

    linkCount = Int(conn.Execute("select count(*) from blog_links")(0))
    Set linkDB = Server.CreateObject("ADODB.RecordSet")
    linkDB.Open "blog_links", Conn, 1, 2
    linkDB.addNew
    linkDB("link_Name") = link_Name
    linkDB("link_URL") = link_URL
    linkDB("link_Image") = link_Image
    linkDB("link_Order") = linkCount
    linkDB("link_IsShow") = False
    linkDB.update
    linkDB.Close
    Set linkDB = Nothing
    showmsg "友情链接添加成功", "<b>网友情链接添加成功,请等待审核！</b><br/><a href=""BlogLink.asp"">返回</a>", "MessageIcon", ""
End If
On Error Resume Next
server.Execute("post/link.html")
If Err Then
    Err.Clear
    Dim blog_Links, ImgLink, TextLink
    Set blog_Links = conn.Execute("select * from blog_Links where link_IsShow=true order by link_Order asc")
    SQLQueryNums = SQLQueryNums + 1
    Do Until blog_Links.EOF
        If Len(blog_Links("link_Image"))>0 Then
            ImgLink = ImgLink&"<a href="""&blog_Links("link_URL")&""" target=""_blank""><img src="""&blog_Links("link_Image")&""" alt="""&blog_Links("link_Name")&""" border=""0"" style=""margin:3px;width:88px;height:31px""/></a>"
        Else
            TextLink = TextLink&"<div class=""link"" style=""width:108px;float:left;overflow:hidden;margin-right:8px;height:24px;line-height:180%""><a href="""&blog_Links("link_URL")&""" target=""_blank"" title="""&blog_Links("link_Name")&""">"&blog_Links("link_Name")&"</a></div>"
        End If
        blog_Links.movenext
    Loop

%>
               <div class="Content">
                 <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
                   <h1 class="ContentTitle"><img src="images/image.gif" alt="" style="margin:0px 2px -3px 0px" class="CateIcon"/><b>图象链接</b></h1>
                   <h2 class="ContentAuthor">Image Link</h2>
                 </div>
                 <div class="Content-body"><%=ImgLink%></div>
               </div>
               <div class="Content">
                 <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
                   <h1 class="ContentTitle"><img src="images/html.gif" alt="" style="margin:0px 2px -3px 0px" class="CateIcon"/><b>文字链接</b></h1>
                   <h2 class="ContentAuthor">Text Link</h2>
                 </div>
                 <div class="Content-body"><%=TextLink%></div>
               </div>
         <%end If%>
         
           <div id="MsgContent" style="width:94%;">
                <div id="MsgHead">申请友情链接</div>
                <div id="MsgBody">
                <form name="frm" action="BlogLink.asp" method="post" onsubmit="return CheckPost()" style="margin:0px;">	  
          	  <table width="100%" cellpadding="0" cellspacing="0">
          	  <tr><td align="right" width="70"><strong>网站名称:</strong></td><td align="left" style="padding:3px;"><input name="link_Name" type="text" size="35" class="userpass" maxlength="40"/> <span style="color:#f00">*</span></td></tr>
          	  <tr><td align="right" width="70"><strong>网站地址:</strong></td><td align="left" style="padding:3px;"><input name="link_URL" type="text" size="50" class="userpass"/> <span style="color:#f00">*</span></td></tr>
          	  <tr><td align="right" width="70"><strong>网站Logo:</strong></td><td align="left" style="padding:3px;"><input name="link_Image" type="text" size="50" class="userpass"/></td></tr>
          	   <tr><td align="right" width="70"><strong>验证码:</strong></td><td align="left" style="padding:3px;"><input name="link_validate" type="text" size="4" class="userpass" maxlength="4" onfocus="this.select()"/> <%=getcode()%></td></tr>
			  <tr><td align="right" width="70"></td><td align="left">提示：网站的Logo和地址要写完整,必须包含 http://</td></tr>
                    <tr>
                      <td colspan="2" align="center" style="padding:3px;">
                        <input name="action" type="hidden" value="postLink"/>
          			    <input type="submit" class="userbutton" value="提交链接"/>
                        <input name="button" type="reset" class="userbutton" value="重写"/>
                        </td>
                    </tr>
          	  </table></form>
          </div></div>

       </div>
       <div id="mainContent-bottomimg"></div>
   </div>
   </div>
   <%Side_Module_Replace '处理系统侧栏模块信息%>
   <div id="sidebar">
	   <div id="innersidebar">
		     <div id="sidebar-topimg"><!--工具条顶部图象--></div>
			  <%=side_html%>
		     <div id="sidebar-bottomimg"></div>
	   </div>
   </div>
   <div style="clear: both;height:1px;overflow:hidden;margin-top:-1px;"></div>
  </div>
 <!--#include file="footer.asp" -->
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
