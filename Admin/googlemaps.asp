
<!--#include file="top.asp"-->
<!--#include file="googlemaps_sub.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
<script language="javascript" src="../jquery/jquery.js"></script>
<TITLE>地图生成</TITLE>
</head>
<body>
  <table width="85%" border="0" cellspacing="0" cellpadding="0"> 
    <tr>
      <td height="30" colspan="2" align="center"></td>
    </tr> 
    <tr>
      <td height="30" colspan="2" align="center">
        <input type="button"  value="点这里生成网站地图XML文档" class="bselect" onClick="createMaps()" ><br><span id="msg" style="color:#F00; display:none">
        <img src="/jquery/images/loading2.gif">正在生成。。。如果网站数据多，则生成的时间会比较长，请不要进行其它的操作，耐心等待生成完成。。</span></td>
    </tr>
  
    <tr>
      <td height="30" colspan="2" align="center">
        Google Sitemaps地址：<a href="<%=WebUrlAddr&GsUrlAddr%>" target="_blank"><font color="#FF0000"><%=GsUrlAddr%></font></a>
        <input type="button" name="Submit2" value="复制" class="badd80 s2tos4" onClick="clipboardData.setData('text','<%=WebUrlAddr&GsUrlAddr%>');alert('复制成功!');"></td>
    </tr>
    <tr>
      <td height="30" colspan="2" align="center">上次生成时间：<font color="#FF0000"><%=ShowDatecreated(server.MapPath(GsUrlAddr))%></font></td>
    </tr>
    <tr>
      <td height="30" colspan="2" align="center">上次文件大小：<%=GetFileSizeStyle(server.MapPath(GsUrlAddr))%></td>
    </tr>
  </table>
<script language="javascript">
function createMaps()
{	
	$.ajax({
	type:"get",	
	dataType:"html",	
	cache:false,
	url:"googlemaps_save.asp",
	beforeSend:function(XMLHttpRequest){
		$("#msg").css("display","");
		},
	success:function(data, textStatus){
		$("#msg").html(data);
		},
	complete: function(XMLHttpRequest, textStatus){
			//HideLoading();
		},
	error: function(){
			alert("生成时产生错误，操作没有完成。");
		}
	
	});
}
</script>
</body></html>

