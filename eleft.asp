<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="22" width="100%" align="center" class="title"><img src="images/yuandian.gif" width="8" height="8" />Categories<img src="images/yuandian.gif" alt="" width="8" height="8" /></td>
      </tr>
      <tr>
        <td height="9">&nbsp;</td>
      </tr>
      <%
	  url=lcase(Request.ServerVariables("Url"))
	  if instr(url,"index.asp")>0 or instr(url,"default.asp")>0 or instr(url,".asp")=0 then snum=10 else snum=50
	  
	  response.Write(getclass(snum,1))%>     
    </table>
   <table width="100%" height="5" border="0" cellspacing="0" cellpadding="0"><tr><td>&nbsp;</td></tr></table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="15" colspan="2" align="center"><img src="images/econtactbg.gif" width="175" height="15" /></td>
        </tr>
        <tr>
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td width="41" height="11" align="right"><img src="images/tel.gif" width="24" height="10" /></td>
          <td width="134" height="11" class="kong2">020-39977900</td>
        </tr>
         <tr>
          <td width="41" height="11" align="right"><img src="images/fax.gif" width="24" height="10" /></td>
          <td width="134" height="11"class="kong2">020-39977900</td>
        </tr>
         <tr>
          <td width="41" height="11" align="right"><img src="images/email.gif" width="24" height="10" /></td>
          <td width="134" height="11"class="kong2">1052867200@qq.com</td>
        </tr>
         <tr>
          <td width="41" height="11" align="right"><img src="images/email.gif" width="24" height="10" /></td>
          <td width="134" height="11"class="kong2">495175304@qq.com</td>
        </tr>
    </table>
    <table width="100%" height="5" border="0" cellspacing="0" cellpadding="0"><tr><td>&nbsp;</td></tr></table>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="60" align="right"><img src="images/wenjuan.gif" width="48" height="41" /></td>
    <td width="115" align="center"><img src="images/sanjiao.gif" width="4" height="5" /><a href="egbook.asp" title="Message" class="xxy">Message</a><a href="javascript:;" class="xxy"></a></td>
  </tr>
  <tr>
    <td align="right"><img src="images/jiaoliu.gif" width="48" height="45" /></td>
    <td align="center"><img src="images/sanjiao.gif" alt="" width="4" height="5" /><a href="econtact.asp" title="Contact US" class="xxy">Contact</a><a href="javascript:;" class="xxy"></a></td>
  </tr>
</table>
