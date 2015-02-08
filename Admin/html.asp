
<!--#include file="top.asp"-->
<%
if cityviewmode=0 then
	response.Write("您的网站是动态浏览模式，不需要静态生成!")
	response.End()
end if
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript" src="/jquery/jquery.js"></script>
<TITLE>备份数据库</TITLE>

</HEAD>
<BODY leftmargin="20">
<table width="98%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> 　
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25"> <div align="left" style="background-image: url('../Images/topbg1.gif')"><strong>生成静态</strong></div>
            </td>
        </tr>
        <tr> 
         
            <td> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="2" height="126">
                <tr> 
                  <td height="25" class="tdbg"><strong>生成静态 需要FSO权限</strong>[<a href="asp.asp">查看是否有FSO权限</a>] <span style="color:#F00">注意：一旦提交，就不能停止，也不能进行其它的操作，请耐心等待生成完成</span></td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"><div align="center"> 
                      <input name="submit" type="button" onClick="createHtml();" value="一键生成">
                    </div></td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"><span id="msg" style="color:#F00; display:none"><img src="/jquery/images/loading2.gif">正在生成。。。如果网站数据多，则生成的时间会比较长，请不要进行其它的操作，耐心等待生成完成。。</span></td>
                </tr>
                <tr> 
                  <td height="47" class="tdbg">　</td>
                </tr>
              </table></td>
         
        </tr>
      </table>
     
    </td>
  </tr>
</table>
<script language="javascript">
function createHtml()
{	
	$.ajax({
	type:"get",	
	dataType:"html",	
	cache:false,
	url:"html_save.asp",
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
			alert("生成时产生错误，操作没有全部完成。");
		}
	
	});
}
</script>
</BODY>
</HTML>
