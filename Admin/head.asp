﻿
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<LINK href="inc/style.css" type=text/css rel=stylesheet>
<title></title>
<style type="text/css">
a:link { color:#000000;text-decoration:none;}
a:hover {color:#666666;text-decoration:none;}
a:visited {color:#000000;text-decoration:none;}
a:active {color: #666666;text-decoration:none;}
TD {FONT-SIZE: 9pt; COLOR: #000000; LINE-HEIGHT: 14pt;}
A {FONT-SIZE: 9pt; COLOR: #blue; LINE-HEIGHT: 14pt; TEXT-DECORATION: none;}
td {FONT-SIZE: 9pt; FILTER: dropshadow(color=#FFFFFF,offx=1,offy=1); COLOR: #000000; FONT-FAMILY: "宋体"}
img {filter:Alpha(opacity:100); chroma(color=#FFFFFF)}
</style>
<base target="main">
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
		parent.frame.cols="0,*";
		displayBar=false;
		obj.src="inc/top_open.gif";
		obj.title="打开左边管理菜单";
	}
	else{
		parent.frame.cols="140,*";
		displayBar=true;
		obj.src="inc/top_close.gif";
		obj.title="关闭左边管理菜单";
	}
}
</script>
</HEAD>
<body leftmargin="0" topmargin="0" bgcolor="#e9e9e9">
<table height="100%" width="100%" border=0 cellpadding=0 cellspacing=0 bgcolor="#e9e9e9">
  <tr valign=middle> 
	<td width=50>
	<img onClick="switchBar(this)" src="inc/top_close.gif" title="关闭左边管理菜单" style="cursor:hand">
	</td>
    <td width=79><a href="right.asp" target=main>管理首页</a></td>
    <td width=79><a href="asp.asp" target="main">服务器变量</a></td>
    <td width=81><a href="exit.asp" target="_parent">退出管理</a></td>
    <td width="85">&nbsp;</td>
    <td width="480">&nbsp;</td>
  </tr>
  <tr bgcolor="#EFEBEF"><td height="13" colspan="6"></td></tr>
</table>
