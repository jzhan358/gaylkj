
<%sub s_body()%>

<HTML>
<HEAD>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>
<style type="text/css">
A{TEXT-DECORATION: none}
A:link {COLOR: #666666; FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:visited {COLOR: #666666; FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:active {FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:hover {BORDER-BOTTOM: 1px dotted; BORDER-LEFT-WIDTH: 1px; BORDER-RIGHT-WIDTH: 1px; BORDER-TOP-WIDTH: 1px; COLOR: #ff6600; TEXT-DECORATION: none}
BODY {
FONT-SIZE: 12px;
COLOR: #666666;
FONT-FAMILY:  宋体;
background-color: #ffffff; 
background-image: url(images/show.gif);
SCROLLBAR-FACE-COLOR: #e8e7e7; 
SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
SCROLLBAR-SHADOW-COLOR: #ffffff; 
SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
SCROLLBAR-ARROW-COLOR: #ff6600; 
SCROLLBAR-TRACK-COLOR: #EFEFEF; 
SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
SCROLLBAR-BASE-COLOR: #000000
}
TABLE {BORDER-COLLAPSE: collapse; FONT-FAMILY: 宋体; FONT-SIZE: 9pt}
.button{height:18px;width:62px;background:#f6f6f9 url(images/ButtonBg.gif); border:solid 1px #5589AA;color: #000000 ;FONT-SIZE: 9pt}
.lanyu{border:solid 1px #5589AA;color: #000000 ; font-size: 12px;}
.font {  filter: DropShadow(Color=#cccccc, OffX=2, OffY=1, Positive=2); text-decoration: none; font-size: 9pt}
</style>
<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>
      <br>
	  
	  <% 
	  sql="select * from s_info "
      set rs=server.createobject("adodb.recordset")
      rs.open sql,conn,1,1
	  if Not(rs.BOF or rs.EOF) then
	   %>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="viewabout" id="viewabout">查看服务宗旨</a></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><%=rs("s_text")%></td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              [<a href="admin_s.asp?action=edits">修改服务宗旨</a>] </td>
          </tr>
      </table>
	  <% 
	  End if
	  rs.close
	  set rs=Nothing
	  
	  if request.QueryString("action")="adds" then
     %>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form9" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="addabout" id="addabout">建立服务宗旨</a></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
              <textarea name="s_text" style="display:none"></textarea>
             <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=s_text&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定增加"> 
              [<a href="admin_s.asp">返回</a>] </td>
          </tr>
          <input type="hidden" name="action" value="adds">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<% End if
   if request.QueryString("action")="edits" then
   	  sql="select * from s_info "
      set rs=server.createobject("adodb.recordset")
      rs.open sql,conn,1,3
	  if Not(rs.BOF or rs.EOF) then
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form9" method="post" action="">
          <tr> 
            <td bgcolor="#E8E8E8">修改服务宗旨</td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
              <textarea name="s_text" style="display:none"><%=Server.HTMLEncode(rs("s_text"))%></textarea>
             <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=s_text&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> 
              [<a href="admin_s.asp">返回</a>] </td>
          </tr>
          <input type="hidden" name="action" value="edits">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
<%    End if
      rs.close
	  set rs=Nothing
      End if
	   %>

      <br>
    </td>
  </tr>
  <tr> 
    <td colspan="3" height="1" background="images/dotlineh.gif"></td>
  </tr>
</table>
<%
End sub
%>