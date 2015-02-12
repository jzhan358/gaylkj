<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--内容-->
 <div id="Tbody">
<%
'==================================
'  用户登录页面
'    更新时间: 2006-5-29
'==================================
If Request.QueryString("action") = "logout" Then
    logout(True)

%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead">退出系统</div>
      <div id="MsgBody">
		 <div class="MessageIcon"></div>
        <div class="MessageText"><b>退出登录成功</b><br/>
         <a href="default.asp">单击返回首页</a></div>
		</div>
	  </div>
	</div>
<br/><br/>
 <%
ElseIf Request.Form("action") = "login" Then
    Dim loginUser
    loginUser = login(Request.Form("UserName"), Request.Form("Password"))
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=loginUser(0)%></div>
      <div id="MsgBody">
	   <div class="<%=loginUser(2)%>"></div>
       <div class="MessageText"><%=loginUser(1)%></div>
	  </div>
	</div>
  </div><br/><br/>
<%
Else
%><br/><br/>
   <div style="text-align:center;">
  <form name="checkUser" action="login.asp" method="post">
    <div id="MsgContent">
      <div id="MsgHead">用户登录</div>
      <div id="MsgBody">
	  <input name="action" type="hidden" value="login"/>
	   <label>用户名：<input name="username" type="text" size="18" class="userpass" maxlength="24"/></label><br/>
	   <label>密　码：<input name="password" type="password" size="18" class="userpass"/></label><br/>
	   <label>验证码：<input name="validate" type="text" size="4" class="userpass" maxlength="4" onfocus="this.select()"/> <%=getcode()%></label><br/>
	   　　<label><input name="KeepLogin" type="checkbox" value="1"/>记住我的登录信息</label><br/>
	   <input type="submit" value="登　陆" class="userbutton"/>　<input type="button" value="用户注册" class="userbutton" onclick="location='register.asp'"/>
	   </div>
	</div>
  </form>
  </div><br/><br/>
 <%End if%>
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
