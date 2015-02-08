
<!--#include file="top.asp"-->
<html>
<head>
<title>管理</title>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
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
		obj.title="打开左边管理菜单";
	}
	else{
		this.frame.cols="140,*";
		displayBar=true;
		obj.src="inc/top_close.gif";
		obj.title="关闭左边管理菜单";
	}
}
</script>
<SCRIPT language=javascript>
<!--

function menu_display(t_id,i_id){//显示隐藏程序
        var t_id;//表格ID
        var i_id;//图片ID
        var on_img="images/fold_2.gif";//打开时图片
        var off_img="images/fold_1.gif";//隐藏时图片
                if (t_id.style.display == "none") {//如果为隐藏状态
                t_id.style.display="";//切换为显示状态
                i_id.src=on_img;//换图
}
        else{//否则
                t_id.style.display="none";//切换为隐藏状态
                i_id.src=off_img;
                }//换图

}

function showhide(t_id,i_id){//显示隐藏程序
        var t_id;//表格ID
        var i_id;//图片ID
        var on_img="images/minus.gif";//打开时图片
        var off_img="images/plus.gif";//隐藏时图片
                if (t_id.style.display == "none") {//如果为隐藏状态
                t_id.style.display="";//切换为显示状态
                i_id.src=on_img;//换图
}
        else{//否则
                t_id.style.display="none";//切换为隐藏状态
                i_id.src=off_img;
                }//换图

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
                  align=absMiddle></a> <STRONG>后台管理</STRONG> <IMG height=9 
                  src="images/now.gif" width=12 align=absMiddle></TD>
    </TR>
    <TR> 
      <TD height="25" bgcolor="#6B696B" class=leftmenu1 
                onclick=javascript:menu_display(table1,img1); width="127"><IMG id=img1 
                  height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">系统管理</span></TD>
    </TR>
  <TBODY id=table1 >
    <!-- 正常情况下  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_config.asp">基本信息</a></TD>
    </TR>
	<TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_config.asp?action=editconfig">修改配置</a></TD>
    </TR>
     <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="html.asp">静态生成</a></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="googlemaps.asp">网站地图生成</a></TD>
    </TR>
    <!-- 正常情况下  -->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="backup.asp">备份数据库</a></TD>
    </TR>
    <!-- 正常情况下  -->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="admin.asp">管理员密码修改</a></TD>
    </TR>
    
    <!-- 正常情况下  -->
  </TBODY>
  <TBODY>
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table2,img2) width="127"><IMG 
                  id=img2 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="leftmenu2 style1">公司信息</span></TD>
    </TR>
  <TBODY id=table2 >
    <!-- 正常情况下  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="about.asp">中文</a> 
      </TD>
    </TR>
     <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="about.asp?language=1">英文</a> 
      </TD>
    </TR>
	
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="about.asp?action=newabout">新建简介</A> </TD>
    </TR>
  </TBODY>
    <TBODY>
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table7,img7) width="127"><IMG 
                  id=img7 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="leftmenu2 style1">信息分类</span></TD>
    </TR>
  <TBODY id=table7 >
    <!-- 正常情况下  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
     <TR> 
      <TD height="18" class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> 
        <a href="newsclass.asp">中文分类管理</a></TD>
    </TR>
    <TR> 
      <TD height="18" class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> 
        <a href="newsclass.asp?language=1">英文分类管理</a></TD>
    </TR>
    <!-- 正常情况下  -->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="newsclass.asp?action=new">新增分类</a></TD>
    </TR>
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table3,img3) width="127"><IMG 
                  id=img3 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">公司动态</span></TD>
  </TR>
  <TBODY id=table3 >
    <!-- 正常情况下  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>	   
   
    <!-- 正常情况下  -->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="news_add.asp?cid=1&action=new">新增中文新闻</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="news_add.asp?cid=7&action=new">新增英文新闻</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="news.asp?cid=1">中文新闻列表</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="news.asp?cid=7">英文新闻列表</A></TD>
    </TR>
    <!-- 正常情况下  -->
  </TBODY>
  <TR> 
    <!-- <TD height=1 width="127"></TD>
  </TR>
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table30,img30) width="127"><IMG 
                  id=img30 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">人才招聘</span></TD>
  </TR>
  <TBODY id=table30 >
    <!-- 正常情况下 
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>    
    <!-- 正常情况下  
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_news.asp?id=2">招聘管理</a></TD>
    </TR>
    <!-- 正常情况下  
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_news.asp?id=2&action=newnews#newnews">新增招聘</A></TD>
    </TR>
    正常情况下  -->
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table40,img40) width="127"><IMG 
                  id=img40 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">技术资料</span></TD>
  </TR>
  <TBODY id=table40 >
    <!-- 正常情况下  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>    
    <!-- 正常情况下  -->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="news_add.asp?cid=11&action=new">新增中文资料</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="news_add.asp?cid=12&action=new">新增英文资料</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="news.asp?cid=11">中文资料管理</a></TD>
    </TR>
    <!-- 正常情况下  -->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="news.asp?cid=12">英文资料管理</a></TD>
    </TR>
    <!-- 正常情况下  -->
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table4,img4) width="127"><IMG 
                  id=img4 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">产品管理</span></TD>
  </TR>
  <TBODY id=table4 >
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="proclass.asp">一级分类管理</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="proclass.asp?action=new">新增一级分类</a> 
      </TD>
    </TR>
	<!--
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="admin_p_small.asp">二级分类管理</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="admin_p_small.asp?action=newp">新增二级分类</a> 
      </TD>
    </TR>-->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <A 
                  class=linkstyle2 
                  href="products.asp">产品管理</A></TD>
    </TR>
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="products_edit.asp?action=new">添加产品 
        </a></TD>
    </TR>
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TR> 
    <TD height="25" bgcolor="#6B696B" class=leftmenu1 onclick=menu_display(table6,img6) width="127"><IMG 
                  id=img6 height=16 src="images/fold_2.gif" width=16 
                  align=absMiddle> <span class="style1">客服中心</span> </TD>
  </TR>
  <TBODY id=table6 >
    <!-- 正常情况下  -->
    <TR> 
      <TD height=1 width="127"></TD>
    </TR>
    <!-- 正常情况下  -->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="book.asp">客户留言</a></TD>
    </TR>
    <!-- 正常情况下  -->
    <TR> 
      <TD class=leftmenu2 width="127">　<IMG src="images/plus.gif" width="9" height="9"> <a href="order.asp">在线订购</a></TD>
    </TR>
    <!-- 正常情况下  -->
    <!-- 正常情况下  -->
    <!-- 正常情况下  -->
  </TBODY>
  <TR> 
    <TD height=1 width="127"></TD>
  </TR>
  <TBODY>
 <TR>
 <TD align="center" width="127"><br>
        版权所有:<%=cityname%></TD>
    </TR>  
  </TBODY>
</TABLE>
</body>
</html>