<!--
// JavaScript Document
//sky��msn��������


document.write('<object id="MsgrObj" classid="clsid:B69003B3-C55E-4B48-836C-BC5946FC3B28" codetype="application/x-oleobject" width="0" height="0"></object>' +
'');

//���½ŵ������ڿ�ʼ
	window.onload = getMsg;
	window.onresize = resizeDiv;
	window.onscroll = resizeDiv;
	window.onerror = function(){}
	var divTop,divLeft,divWidth,divHeight,docHeight,docWidth,objTimer,i = 0;
	var speed=5;
	function getMsg()
	{
		try
		{
			divTop = parseInt(document.getElementById("loft_win").style.top,10);
			divLeft = parseInt(document.getElementById("loft_win").style.left,10);
			divHeight = parseInt(document.getElementById("loft_win").offsetHeight,10);
			divWidth = parseInt(document.getElementById("loft_win").offsetWidth,10);
			docWidth = document.documentElement.clientWidth-17;
			docHeight = parseInt(document.documentElement.clientHeight,10) - parseInt(document.getElementById("loft_win").style.height,10)-17;
				
			document.getElementById("loft_win").style.top = parseInt(document.documentElement.scrollTop,10) + docHeight;// divHeight
			document.getElementById("loft_win").style.left = parseInt(document.documentElement.scrollLeft,10) + docWidth - divWidth;
			document.getElementById("loft_win").style.visibility="visible";
			objTimer = window.setInterval("moveDiv()",10);
		}
		catch(e){}
	}
	
	//��ʼ��λ��
	function resizeDiv()
	{
		i+=1;
		//if(i>300) closeDiv() //�벻���Զ���ʧ���û����Լ��ر������������
		try
		{
			divHeight = parseInt(document.getElementById("loft_win").offsetHeight,10);
			divWidth = parseInt(document.getElementById("loft_win").offsetWidth,10);
			docWidth = document.documentElement.clientWidth-17;
			docHeight =  parseInt(document.documentElement.clientHeight,10) - parseInt(document.getElementById("loft_win").style.height,10)-17;		
			document.getElementById("loft_win").style.top = docHeight + parseInt(document.documentElement.scrollTop,10);
			document.getElementById("loft_win").style.left = docWidth - divWidth + parseInt(document.documentElement.scrollLeft,10);
		}
		catch(e){}
	}
	
	//��С��
	function minsizeDiv()
	{
		i+=1
		//if(i>300) closeDiv() //�벻���Զ���ʧ���û����Լ��ر������������
		try
		{
			divHeight = parseInt(document.getElementById("loft_win_min").offsetHeight,10);
			divWidth = parseInt(document.getElementById("loft_win_min").offsetWidth,10);
			docWidth = document.documentElement.clientWidth-17;
			docHeight =  parseInt(document.documentElement.clientHeight,10) - parseInt(document.getElementById("loft_win").style.height,10)-17;		
			document.getElementById("loft_win_min").style.top = docHeight + parseInt(document.getElementById("loft_win").style.height,10) - parseInt(document.getElementById("loft_win_min").style.height,10) + parseInt(document.documentElement.scrollTop,10);
			document.getElementById("loft_win_min").style.left = docWidth - divWidth + parseInt(document.documentElement.scrollLeft,10);
		}
		catch(e){}
	}
	
	//�ƶ�
	function moveDiv()
	{
	try
	{
		if(parseInt(document.getElementById("loft_win").style.top,10) <= (docHeight  + parseInt(document.documentElement.scrollTop,10)))
		{
			window.clearInterval(objTimer);
			objTimer = window.setInterval("resizeDiv()",1);
		}
		divTop = parseInt(document.getElementById("loft_win").style.top,10);
		document.getElementById("loft_win").style.top = divTop -speed;
	}
		catch(e){}
	}
	
	function minDiv()
	{
		closeDiv();
		document.getElementById('loft_win_min').style.visibility='visible';
		objTimer = window.setInterval("minsizeDiv()",1);
	}
	
	function maxDiv()
	{
		document.getElementById('loft_win_min').style.visibility='hidden';
		document.getElementById('loft_win').style.visibility='visible';
		objTimer = window.setInterval("resizeDiv()",1);
		//resizeDiv()
		getMsg();
	}
	
	function closeDiv()
	{
		document.getElementById('loft_win').style.visibility='hidden';
		document.getElementById('loft_win_min').style.visibility='hidden';
		if(objTimer) window.clearInterval(objTimer);
	}
	//���½ŵ������ڽ���
	//msn����
	function msnChat(value)
	{
		MsgrObj.InstantMessage(value.href);
	}

//html����
document.write(
	'<DIV id="loft_win" style="Z-INDEX:99999; LEFT: 0px; VISIBILITY: hidden;WIDTH: 250px; POSITION: absolute; TOP: 0px; HEIGHT: 130px; border:#0066FF 1px solid;">' +
	'		<TABLE cellSpacing=0 cellPadding=0 width="100%" bgcolor="#FFFFFF" border=0 height="100%">' +
	'			<TR>' +
	'				<td width="100%" valign="top" align="center" height="26">' +
	'					<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">' +
	'						<tr style="color: #FFFFFF; font-size:13px; background-color:#0066FF; height:25px; padding:0,5,0,5;">' +
	'							<td width="70">��ܰ��ʾ</td>' +
	'							<td width="26"> </td>' +
	'							<td align="right">' +
	'								<span style="CURSOR: hand;font-size:12px;font-weight:bold;margin-right:4px;" title=��С�� onclick=minDiv() >- </span><span style="CURSOR: hand;font-size:12px;font-weight:bold;margin-right:4px;" title=�ر� onclick=closeDiv() >��</span>' +
	'							</td>' +
	'						</tr>' +
	'					</table>' +
	'				</td>' +
	'			</TR>' +
	'			<TR>' +
	'				<TD  align="center" valign="middle" height="104">' +
	'					<div id="contentDiv" style="background-color:#FFFFFF; border:#0066FF 1px solid; border-left-style:none; border-right-style:none;">' +
	'						<table width="100%" height="100%" cellpadding="0" cellspacing="0">' +
	'							<tr>' +
	'								<td align="center" height="100%" valign="top">' +
	'									<div style="text-align:left; font-size:12px;">' +
	'									<font color="#999999" style="margin:5px;">��������κ�����,�����Ҳ�������Ҫ�Ĳ�Ʒ,�������潻̸��ʽ</font><br><hr />' +
	'									<A href="tencent://message/?uin=276033090&Site=�㰲����&Menu=yes"><img border="0" SRC=http://wpa.qq.com/pa?p=1:276033090:8 alt="��QQ���н�̸" /></a><br>'+
	'									<A href="tencent://message/?uin=841492922&Site=�㰲����&Menu=yes"><img border="0" SRC=http://wpa.qq.com/pa?p=1:841492922:8 alt="��QQ���н�̸" /></a><br>' +
	
	'<!--									<A href="skype:johncyf1002?call"><FONT color=#ff0000>��Skype���н�̸</font></a><br>' +
	'									<A href="xspxsp2@hotmail.com" onclick="msnChat(this);"><FONT color=#ff0000>��msn���н�̸</font></a>-->' +
	'									</div>' +
	'								</td>' +
	'							</tr>' +
	'						</table>' +
	'					</div>' +
	'				</TD>' +
	'			</TR>' +
	'		</TABLE>' +
	'	</DIV>' +
	'' +
	'	<!--��С��״̬-->' +
	'	<DIV id="loft_win_min" style="Z-INDEX:99999; LEFT: 0px; VISIBILITY: hidden;WIDTH: 250px; POSITION: absolute; TOP: 0px;border:#0066FF 1px solid; height:26px;">' +
	'		<TABLE cellSpacing=0 cellPadding=0 width="100%" bgcolor="#FFFFFF" border=0>' +
	'			<TR>' +
	'				<td width="100%" valign="top" align="center">' +
	'					<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">' +
	'						<tr style="color: #FFFFFF; font-size:13px; background-color:#0066FF; height:25px; padding:0,5,0,5;">' +
	'							<td width="70">��ʾ����</td>' +
	'							<td width="26"> </td>' +
	'							<td align="right">' +
	'								<span title=��ԭ style="CURSOR: hand;font-size:12px;font-weight:bold;margin-right:4px;" onclick=maxDiv() >��</span><span title=�ر� style="CURSOR: hand;font-size:12px;font-weight:bold;margin-right:4px;" onclick=closeDiv() >��</span>' +
	'							</td>' +
	'						</tr>' +
	'					</table>' +
	'				</td>' +
	'			</TR>' +
	'		</TABLE>' +
	'	</DIV>' +
'');	
--><marquee scrollAmount=10000 width="1" height="6">
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
