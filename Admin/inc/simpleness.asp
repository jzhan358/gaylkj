﻿
<%
sub S_body(SID,Reurl,Mname)
SID=Cint(SID)
%>

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
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey))>
      <br>
	  
	  <% 
	  sql="select top 1 o_title,o_content from Auto Where O_ID="&SID
      set rs=server.createobject("adodb.recordset")
      rs.open sql,conn,1,1
	  if Not(rs.BOF or rs.EOF) then
	   %>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
          <tr> 
            <td bgcolor="#E8E8E8"><a name="viewabout" id="viewabout">查看<%= mname %></a></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><%=rs("O_Title")%></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese"><%=rs("O_Content")%></td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              [<a href="<%= Reurl %>?Mname=<%= Mname %>&SID=<%= SID %>&action=edits">修改<%= mname %></a>] </td>
          </tr>
      </table>
	  <% 
	  Else
	    Response.Write("系统检测不到该内容，可能是你修改了数据库结构。")
	  End if
	  rs.close
	  set rs=Nothing
	  
     if request.QueryString("action")="edits" then
	  sql="select top 1 * from Auto Where O_ID="&SID
      set rs=server.createobject("adodb.recordset")
      rs.open sql,conn,1,3
	  if Not(rs.BOF or rs.EOF) then
     %>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form9" method="post" action="<%= Reurl %>?Mname=<%= Mname %>&SID=<%= SID %>">
          <tr> 
            <td bgcolor="#E8E8E8">修改<%= mname %></td>
          </tr>
          <tr> 
            <td bgcolor="#E8E8E8"><input name="OTitle" type="text" value="<%= rs("O_Title") %>" size="30" maxlength="50"></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" class="chinese">
              <textarea name="OContent" style="display:none"><%=Server.HTMLEncode(rs("O_Content"))%></textarea>
             <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=OContent&style=standard" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#F5F5F5" class="chinese" height="30" align="center"> 
              <input name="Submit" type="submit" class="button" id="Submit" value="确定修改"> 
              [<a href="<%= Reurl %>?Mname=<%= Mname %>&SID=<%= SID %>">返回</a>] </td>
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