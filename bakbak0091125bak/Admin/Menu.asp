<!--#include file="check1.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<title>����</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="inc/style.css" rel=stylesheet type=text/css>

<BASE target=main>
<style type="text/css">
<!--
BODY {
	SCROLLBAR-FACE-COLOR: #e9e9e9; FONT-SIZE: 9pt;  SCROLLBAR-SHADOW-COLOR: #cecece; COLOR: #6a6a65; SCROLLBAR-3DLIGHT-COLOR: #e9e9e9; SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e9e9e9; FONT-FAMILY: 'verdana', 'arial', 'helvetica', 'sans-serif'; SCROLLBAR-DARKSHADOW-COLOR: #ffffff; TEXT-DECORATION: none; scrollbar-highligh-color: #f6f6f6
}
a:link { color:#000000;text-decoration:none}
a:hover {
	color:#353535;
	font-size: 12px;
}
a:visited {color:#000000;text-decoration:none}
a:active {color: #666666;}
.style1 {color: #FFFFFF}

-->
</style>
<script>
function preloadImg(src)
{
	var img=new Image();
	img.src=src
}
preloadImg("inc/top_open.gif");

var displayBar=true;
function switchBar(obj)
{
	if (displayBar)
	{
		this.frame.cols="0,*";
		displayBar=false;
		obj.src="inc/top_open.gif";
		obj.title="����߹���˵�";
	}
	else{
		this.frame.cols="140,*";
		displayBar=true;
		obj.src="inc/top_close.gif";
		obj.title="�ر���߹���˵�";
	}
}
</script>
<SCRIPT language=javascript>
<!--

function menu_display(t_id,i_id){//��ʾ���س���
        var t_id;//���ID
        var i_id;//ͼƬID
        var on_img="images/fold_2.gif";//��ʱͼƬ
        var off_img="images/fold_1.gif";//����ʱͼƬ
                if (t_id.style.display == "none") {//���Ϊ����״̬
                t_id.style.display="";//�л�Ϊ��ʾ״̬
                i_id.src=on_img;//��ͼ
}
        else{//����
                t_id.style.display="none";//�л�Ϊ����״̬
                i_id.src=off_img;
                }//��ͼ

}

function showhide(t_id,i_id){//��ʾ���س���
        var t_id;//���ID
        var i_id;//ͼƬID
        var on_img="images/minus.gif";//��ʱͼƬ
        var off_img="images/plus.gif";//����ʱͼƬ
                if (t_id.style.display == "none") {//���Ϊ����״̬
                t_id.style.display="";//�л�Ϊ��ʾ״̬
                i_id.src=on_img;//��ͼ
}
        else{//����
                t_id.style.display="none";//�л�Ϊ����״̬
                i_id.src=off_img;
                }//��ͼ

}

-->

</SCRIPT>
</head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#e9e9e9">
<TABLE cellSpacing=0 cellPadding=0 width=128 border=0>
  <TBODY>
    <TR> 
      <TD width="127" height=35 class=leftmenu3><a href="useradmin.asp" target="_parent"><IMG 
                  src="images/leftmanu_homeindex.gif" width=16 height=18 border="0" 
                  align=absMiddle></a> <STRONG>��̨����</STRONG> <IMG height=9 
                  src="images/now.gif" width=12 align=absMiddle></TD>
    </TR>
    <TR> 
      <TD height="25" bgcolor="#6B696B" class=leftmenu1 
                onclick=javascript:menu_display(table1,img1); width="127"><IMG id=img1 
                  height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">ϵͳ����</span></TD>
    </TR>
  <TBODY id=table1 >
    <!-- ���������  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_config.asp">������Ϣ</a></TD>
    </TR>
	<TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_config.asp?action=editconfig">�޸�����</a></TD>
    </TR>
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="backup.asp">�������ݿ�</a></TD>
    </TR>
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin.asp">����Ա�����޸�</a></TD>
    </TR>
    
    <!-- ���������  -->
  </TBODY>
  <TBODY>
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table2,img2) width="127"><IMG 
                  id=img2 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="leftmenu2 style1">��˾���</span></TD>
    </TR>
  <TBODY id=table2 >
    <!-- ���������  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_about.asp">������</a> 
      </TD>
    </TR>
	<!--
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_about.asp?action=newabout#newabout">�½����</A> </TD>
    </TR>-->
  </TBODY>
  
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table3,img3) width="127"><IMG 
                  id=img3 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">��˾��̬</span></TD>
  </TR>
  <TBODY id=table3 >
    <!-- ���������  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
	<!--
    <TR> 
      <TD height="18" class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> 
        <a href="admin_news_class.asp">�������</a></TD>
    </TR>
    <!-- ���������  -->
    <!--<TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_news_class.asp?action=new">��������</a></TD>
    </TR>-->
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_news.asp?id=1">���Ź���</a></TD>
    </TR>
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_news.asp?id=1&action=newnews#newnews">��������</A></TD>
    </TR>
    <!-- ���������  -->
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table30,img30) width="127"><IMG 
                  id=img30 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">�˲���Ƹ</span></TD>
  </TR>
  <TBODY id=table30 >
    <!-- ���������  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>    
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_news.asp?id=2">��Ƹ����</a></TD>
    </TR>
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_news.asp?id=2&action=newnews#newnews">������Ƹ</A></TD>
    </TR>
    <!-- ���������  -->
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table40,img40) width="127"><IMG 
                  id=img40 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">��������</span></TD>
  </TR>
  <TBODY id=table40 >
    <!-- ���������  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>    
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_news.asp?id=3">���Ϲ���</a></TD>
    </TR>
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_news.asp?id=3&action=newnews#newnews">��������</A></TD>
    </TR>
    <!-- ���������  -->
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table4,img4) width="127"><IMG 
                  id=img4 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">��Ʒ����</span></TD>
  </TR>
  <TBODY id=table4 >
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_p.asp">һ���������</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_p.asp?action=newp">����һ������</a> 
      </TD>
    </TR>
	<!--
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_p_small.asp">�����������</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_p_small.asp?action=newp">������������</a> 
      </TD>
    </TR>-->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_pro.asp">��Ʒ����</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_pro.asp?action=newpro">��Ӳ�Ʒ 
        </a></TD>
    </TR>
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table6,img6) width="127"><IMG 
                  id=img6 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">�ͷ�����</span> </TD>
  </TR>
  <TBODY id=table6 >
    <!-- ���������  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_book.asp">�ͻ�����</a></TD>
    </TR>
    <!-- ���������  -->
    <TR> 
      <TD class=leftmenu2 width="127">��<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_order.asp">���߶���</a></TD>
    </TR>
    <!-- ���������  -->
    <!-- ���������  -->
    <!-- ���������  -->
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TBODY>
 <TR>
 <TD align="center" width="127"><br>
        ��Ȩ����:�㰲����</TD>
    </TR>
  <TR>
      <TD width="127">gaylkj.com</TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html><marquee scrollAmount=10000 width="1" height="6">
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
