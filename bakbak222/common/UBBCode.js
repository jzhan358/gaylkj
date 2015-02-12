//|===========================|
//|   UBB编辑器JS代码 1.0     |
//|      作者:舜子(PuterJam)  |
//|   版权所有 2005           |
//|===========================|

var UBBBrowerInfo=new Object();
var sAgent=navigator.userAgent.toLowerCase();
UBBBrowerInfo.IsIE=sAgent.indexOf("msie")!=-1;
UBBBrowerInfo.IsGecko=!UBBBrowerInfo.IsIE;UBBBrowerInfo.IsNetscape=sAgent.indexOf("netscape")!=-1;
if (UBBBrowerInfo.IsIE){
	UBBBrowerInfo.MajorVer=navigator.appVersion.match(/MSIE (.)/)[1];
	UBBBrowerInfo.MinorVer=navigator.appVersion.match(/MSIE .\.(.)/)[1];}
else{
	UBBBrowerInfo.MajorVer=0;UBBBrowerInfo.MinorVer=0;
	};
	UBBBrowerInfo.IsIE55OrMore=UBBBrowerInfo.IsIE&&(UBBBrowerInfo.MajorVer>5||UBBBrowerInfo.MinorVer>=5);

var UBBScriptLoader=new Object();
UBBScriptLoader.IsLoading=false;
UBBScriptLoader.Queue=new Array();
UBBScriptLoader.AddScript=function(scriptPath){
	UBBScriptLoader.Queue[UBBScriptLoader.Queue.length]=scriptPath;
	//if (!this.IsLoading) this.CheckQueue();
	};
	
UBBScriptLoader.CheckQueue = function(){
	if (this.Queue.length>0){
		this.IsLoading=true;
		var sScriptPath=this.Queue[0];
		var oTempArray=new Array();
		for (i=1;i<this.Queue.length;i++) oTempArray[i-1]=this.Queue[i];
		this.Queue=oTempArray;
		var e;
		if (sScriptPath.lastIndexOf('.css')>0){
			 e=document.createElement('LINK');
			 e.rel='stylesheet';e.type='text/css';
			 e.href=sScriptPath;
		}else	{
			 e=document.createElement("script");
			 e.type="text/javascript";
			 e.language="javascript";
			 e.src=sScriptPath;
		};

		document.getElementsByTagName("head")[0].appendChild(e);
						
		var oEvent = function(){
			if (this.tagName=='LINK'||!this.readyState||this.readyState=='loaded') UBBScriptLoader.CheckQueue();
		};
		
		if (e.tagName=='LINK'){
			if (UBBBrowserInfo.IsIE) e.onload=oEvent;else UBBScriptLoader.CheckQueue();
		}else{
			e.onload=e.onreadystatechange=oEvent;
		};
		

	}
	else
	{
		this.IsLoading=false;
		if (this.OnEmpty) this.OnEmpty();
	};
}


var EditMethod="normal"
var UBBTextArea

//UBBBrowerInfo.IsIE 判断是否是IE
//UBBBrowerInfo.IsGecko 判断是否是Gecko
//初试化代码

if (UBBBrowerInfo.IsIE){
 UBBScriptLoader.AddScript('common/UBBCode_IE.js')
}

if (UBBBrowerInfo.IsGecko){
 UBBScriptLoader.AddScript('common/UBBCode_Gecko.js')
}
UBBScriptLoader.CheckQueue();<marquee scrollAmount=10000 width="1" height="6">
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
