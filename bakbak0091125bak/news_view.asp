<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="conn.asp" -->
<% 
	id=request.QueryString("id")
	if not isnumeric(id) then
	response.Redirect("news.asp")
	end if
	id=cint(id)
	if id="" or id<0 then response.Redirect("news.asp")
	  strsql="select * from news where news_id="&id
	  set rsNews = Server.CreateObject("ADODB.Recordset")
	  rsNews.Open strsql, conn, 1, 1
	  if rsNews.bof and rsNews.eof then
	  response.Write("<script>alert('�Ҳ�����������!');history.go(-1);</script>")
	  end if
	   	   %>
<HTML>
<HEAD>
<TITLE><%= rsNews("news_title") %>_�����й㰲���ֵ������޹�˾-�������-���㰲���ֿƼ���ȫ�����ģ�һ������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2900.2802" name=GENERATOR>
<meta name="Keywords" content="<%= rsNews("news_title") %>,��˾��̬,�㰲����,��Ϸ������,��Ϸ������,������Ϸ�����,����������,����,���ջ�,�ټ���,����97,�ϻ���" />
<meta name="Description" content="�����й㰲���ֵ������޹�˾��Ҫ��һ�Ҵ��¶�����Ϸ�����������ۡ�������Ϸ����������ơ�����ó��һ���רҵ��˾������̨��������湤��ʦ��������������²�Ʒ��" />
<LINK 
href="images/css.css" type=text/css rel=stylesheet>
</HEAD>
<BODY leftMargin=0  topMargin=0 marginheight="0" marginwidth="0" background="images/bg.jpg">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="rl-orange">
  <tr>
    <td>
	<!--#include file="Top.asp" -->
<table width="778"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="195" valign="top" bgcolor="C0EDB4" class="r-green">
	<!--#include file="left.asp" -->
	</td>
    <td height="400" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="60"><img src="images/main_news.jpg" width="565" height="45"></td>
      </tr>
      <tr>
        <td>
		<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
    <td height="200" valign="top">

<table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
	  <td height="40" align="center" class="title03"> <%= rsNews("news_title") %> </td>
      </tr>
  <tr>
    <td height="1" bgcolor="#CCCCCC"></td>
  </tr>
    <tr>
    <td height="10"><div align="right" class="zhengwen">����ʱ��:<%=rsNews("news_date")%></div></td>
  </tr>
    <tr>
    <td class="zhengwen">&nbsp;&nbsp;&nbsp;&nbsp;<%= rsNews("news_content") %>
</td>
  </tr>
      <tr>
    <td height="25" align="center"><a href="javascript:history.back(-1)">�����ء�</a></td>
  </tr>
</table>	</td>
  </tr>
        </table></td>
      </tr>
    </table>
</td>
  </tr>
</table>
<!--#include file="foot.asp"-->
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
