<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/function.asp" -->
<%
'==================================
'  Google SiteMap for PJBlog
'  更新时间: 2005-12-1
'  FlowJZH@gmail.com
'  www.Aryl.net
'==================================

Sub Escape(ByRef s)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "<", "&gt;")
    s = Replace(s, ">", "&lt;")
End Sub

'读取Blog设置信息
getInfo(1)

Response.Charset = "UTF-8"
Response.ContentType = "text/xml"
Response.Expires = 60

Dim cate_ID, FeedRows
cate_ID = CheckStr(Request.QueryString("cateID"))
If IsInteger(cate_ID) = False Then
    SQL = "SELECT top 10 L.log_ID,L.log_PostTime FROM blog_Content AS L,blog_Category AS C WHERE C.cate_ID=L.log_cateID AND L.log_IsDraft=false and C.cate_Secret=false ORDER BY log_PostTime DESC"
Else
    SQL = "SELECT top 10 L.log_ID,L.log_PostTime FROM blog_Content AS L,blog_Category AS C WHERE log_cateID="&cate_ID&" AND C.cate_ID=L.log_cateID AND L.log_IsDraft=false and C.cate_Secret=false ORDER BY log_PostTime DESC"
End If

Dim RS, i
Set RS = Conn.Execute(SQL)
If RS.EOF Then
    ReDim FeedRows(0, 0)
Else
    FeedRows = RS.getrows()
End If
RS.Close
Set RS = Nothing
Call CloseDB
%><?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.google.com/schemas/sitemap/0.84">
  <url>
    <loc><%=SiteURL%></loc>
    <lastmod><%=ISO8601(DateAdd("h",-1,Now))%></lastmod>
    <changefreq>always</changefreq>
    <priority>0.9</priority>
  </url>
<%
If UBound(FeedRows, 1)>0 Then
    Dim iPrior, dtNow

    dtNow = Now

    With Response
        For i = 0 To UBound(FeedRows, 2)
            iPrior = 0.5
            .Write "  <url>"
            .Write vbCrLf
            .Write "    <loc>"
            
             If blog_postFile = 2 Then 
             	.Write SiteURL&"article/"&FeedRows(0, i)&".htm"
            else
          	 	.Write SiteURL&"article.asp?id="&FeedRows(0, i)
            end if
            
            .Write "</loc>"
            .Write vbCrLf
            .Write "    <lastmod>"
            .Write ISO8601(FeedRows(1, i))
            .Write "</lastmod>"
            .Write vbCrLf
            .Write "    <changefreq>"
            If DateDiff("h", FeedRows(1, i), dtNow) < 24 Then
                .Write "hourly"
                iPrior = 0.8
            ElseIf DateDiff("d", FeedRows(1, i), dtNow) < 7 Then
                .Write "daily"
                iPrior = 0.7
            ElseIf DateDiff("ww", FeedRows(1, i), dtNow) < 4 Then
                .Write "weekly"
            ElseIf DateDiff("m", FeedRows(1, i), dtNow) < 12 Then
                .Write "monthly"
            Else
                .Write "yearly"
            End If
            .Write "</changefreq>"
            .Write vbCrLf
            If iPrior <> 0.5 Then
                .Write "    <priority>"
                .Write iPrior
                .Write "</priority>"
                .Write vbCrLf
            End If
            .Write "  </url>"
            .Write vbCrLf
        Next
    End With
End If

Function ISO8601(DateTime)
    Dim DateMonth, DateDay, DateHour, DateMinute, DateWeek, DateSecond

    DateTime = DateAdd("h", -8, DateTime)
    DateMonth = Month(DateTime)
    DateDay = Day(DateTime)
    DateHour = Hour(DateTime)
    DateMinute = Minute(DateTime)
    DateWeek = Weekday(DateTime)
    DateSecond = Second(DateTime)
    If Len(DateMonth)<2 Then DateMonth = "0"&DateMonth
    If Len(DateDay)<2 Then DateDay = "0"&DateDay
    If Len(DateMinute)<2 Then DateMinute = "0"&DateMinute
    If Len(DateHour)<2 Then DateHour = "0"&DateHour
    If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
    ISO8601 = Year(DateTime)&"-"&DateMonth&"-"&DateDay&"T"&DateHour&":"&DateMinute&":"&DateSecond&"Z"
End Function
%>
</urlset>
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
