<%@ Language="VBScript" %>
<%' Option Explicit %>
<!--#include file="conn.asp"-->
<!--#include file="check1.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�������ݿ�</TITLE>
<style>
<!--
BODY
{
	FONT-FAMILY: ����;
	FONT-SIZE: 9pt
}
TD
{
	FONT-SIZE: 9pt
}
A
{
	COLOR: #000000;
	TEXT-DECORATION: none
}
A:hover
{
	COLOR: #9FA9B3;
	TEXT-DECORATION: underline
}
.input
{
	BORDER: #9FA9B3 1px solid;
	FONT-SIZE: 9pt;
	BACKGROUND-color: #E3E3E3}
.backs
{
	BACKGROUND-COLOR: #E3E3E3;
	COLOR: #ffffff;

}
.backq
{
	BACKGROUND-COLOR: #E3E3E3
}
.backc
{
	BACKGROUND-COLOR: #E3E3E3;
	BORDER: medium none;
	COLOR: #999999;
	HEIGHT: 18px;
	font-size: 9pt
}
.fonts
{
	COLOR: #FFFFFF
}
.style1 {BACKGROUND-COLOR: #E3E3E3; color: #130700; }
-->
</STYLE>
</HEAD>
<BODY leftmargin="20">
<table width="98%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> ��
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25"> <div align="left" style="background-image: url('../Images/topbg1.gif')"><strong>�������ݿ�</strong></div>
            <div align="left">
              <%
if request("action")="Backup" then
call backupdata()
else
%>
            </div></td>
        </tr>
        <tr> 
          <form method="post" action="backup.asp?action=Backup">
            <td> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="2" height="333">
                <tr> 
                  <td height="25" class="tdbg"><strong>���������ļ�</strong>[<a href="asp.asp">��ҪFSOȨ��</a>]</td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"> ��ǰ���ݿ�·��</td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"><input type=text size=50 name=DBpath value="<%=db%>"></td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"> �������ݿ�Ŀ¼[��Ŀ¼�����ڣ������Զ�����]</td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"><input type=text size=50 name=bkfolder value=Databackup></td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg">�������ݿ�����[�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����]</td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"> <input type=text size=30 name=bkDBname value=backup.mdb></td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"><div align="center"> 
                      <input name="submit" type=submit value="ȷ��">
                    </div></td>
                </tr>
                <tr> 
                  <td height="22" class="tdbg"><br> <br>
                    �������Ĭ�����ݿ��ļ�Ϊ<%=db%><br>
                    ������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
                    ע�⣺����·��������������ռ��Ŀ¼�����·��</td>
                </tr>
                <tr> 
                  <td height="47" class="tdbg">��</td>
                </tr>
              </table></td>
          </form>
        </tr>
      </table>
      <%end if%>
      <% 
sub backupdata() 
response.Write("Ϊ��ȫ���,�˹����Ѿ��ر�,�������Ա��ϵ!")
'Dbpath=request.form("Dbpath") 
'Dbpath=server.mappath(Dbpath) 
'bkfolder=request.form("bkfolder") 
'bkdbname=request.form("bkdbname") 
'Set Fso=server.createobject("scripting.filesystemobject") 
'if fso.fileexists(dbpath) then 
'If CheckDir(bkfolder) = True Then 
'fso.copyfile dbpath,bkfolder& "\"& bkdbname 
'else 
'MakeNewsDir bkfolder 
'fso.copyfile dbpath,bkfolder& "\"& bkdbname 
'end if 
'response.write "�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" &bkfolder& "\"& bkdbname 
'Else 
'response.write "�Ҳ���������Ҫ���ݵ��ļ���" 
'End if 
'end sub 
''------------------���ĳһĿ¼�Ƿ����------------------- 
'Function CheckDir(FolderPath) 
'folderpath=Server.MapPath(".")&"\"&folderpath 
'Set fso1 = CreateObject("Scripting.FileSystemObject") 
'If fso1.FolderExists(FolderPath) then 
''���� 
'CheckDir = True 
'Else 
''������ 
'CheckDir = False 
'End if 
'Set fso1 = nothing 
'End Function 
''-------------����ָ����������Ŀ¼--------- 
'Function MakeNewsDir(foldername) 
'Set fso1 = CreateObject("Scripting.FileSystemObject") 
'Set f = fso1.CreateFolder(foldername) 
'MakeNewsDir = True 
'Set fso1 = nothing 
'End Function 
%>
    </td>
  </tr>
</table>
</BODY>
</HTML>
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
