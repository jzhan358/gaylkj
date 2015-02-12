<%'---- ASPCode For GuestBookForPJBlog ----%>
<%'---- ASPCode For GuestBookForPJBlogSubItem1 ----%>

<%
        function NewMessage(ByVal action)
             Dim blog_Message
             IF Not IsArray(Application(CookieName&"_blog_Message")) or action=2 Then
             	Dim book_Messages,book_Message
             	Set book_Messages=Conn.Execute("SELECT top 10 * FROM blog_book order by book_PostTime Desc")
             	SQLQueryNums=SQLQueryNums+1
             	TempVar=""
             	Do While Not book_Messages.EOF
             	    if book_Messages("book_HiddenReply") then
                        book_Message=book_Message&TempVar&book_Messages("book_ID")&"|,|"&book_Messages("book_Messager")&"|,|"&book_Messages("book_PostTime")&"|,|"&"[隐藏留言]"
             	     else
                        book_Message=book_Message&TempVar&book_Messages("book_ID")&"|,|"&book_Messages("book_Messager")&"|,|"&book_Messages("book_PostTime")&"|,|"&book_Messages("book_Content")
             	    end if
             		TempVar="|$|"
             		book_Messages.MoveNext
             	Loop
             	Set book_Messages=Nothing
             	blog_Message=Split(book_Message,"|$|")
             	Application.Lock
             	Application(CookieName&"_blog_Message")=blog_Message
             	Application.UnLock
             Else
             	blog_Message=Application(CookieName&"_blog_Message")
             End IF
             
             if action<>2 then
              dim Message_Items,Message_Item
             	For Each Message_Items IN blog_Message
             	 Message_Item=Split(Message_Items,"|,|")
             	 NewMessage=NewMessage&"<a class=""sideA"" href=""LoadMod.asp?plugins=GuestBookForPJBlog#book_"&Message_Item(0)&""" title="""&Message_Item(1)&" 于 "&Message_Item(2)&" 发表留言"&CHR(10)&CCEncode(CutStr(Message_Item(3),25))&""">"&CCEncode(CutStr(Message_Item(3),25))&"</a>"
             	Next
              end if
       end function
   
       '处理最新留言内容
        Dim Message_code
        if Session(CookieName&"_LastDo")="DelMessage" or Session(CookieName&"_LastDo")="AddMessage" then NewMessage(2)
    	Message_code=NewMessage(0)
        side_html_default=replace(side_html_default,"<$NewMsg$>",Message_code)
        side_html=replace(side_html,"<$NewMsg$>",Message_code)
    %>
     
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
