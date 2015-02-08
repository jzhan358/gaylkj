
<%sub diserror()%>
<HTML><HEAD><TITLE>错误提示</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8"><LINK 
href="inc/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="60%" border="0" cellspacing="1" cellpadding="10" align="center" bgcolor="#293863">
  <tr> 
    <td bgcolor="#E8E8E8" class="chinese"><div align="center">出错拉！</div></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" class="chinese"><%=errmsg%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" class="chinese">[<a href="javascript:history.go(-1)">返回</a>]</td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
<%End sub%>