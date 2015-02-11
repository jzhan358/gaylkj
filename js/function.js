//2009.9.27 加入chkSelect函数 2009.10.4 修改chkSelect函数 修改了doPrint函数 2009.11.01修改CheckAll函数
//2009.05.20加入setHome函数 2009.7.23 加入sendData()函数，重新修改本函数库，修改为类的形式 并修改所有函数为参数不限定
//2009.04.06 修改GetCookies有时不能获取值的问题 2009.04.17加入boxs函数 修改了GetCookies的错误
//2009.03.31加入 CheckAll(),showmsg函数 合并了cookies.js 增加了函数说明，修改了chkfocusValue,chkblurValue的通用性
//by selectersky
cjx={
	//设为首页
	setHome:function()
	{
		var obj=arguments[0];
		var url=arguments[1];
		try{obj.style.behavior='url(#default#homepage)';obj.setHomePage(url);}
		catch(e){
			if(window.netscape) {
				try {netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");}
				catch (e) {alert("此操作被浏览器拒绝！\n请在浏览器地址栏输入“about:config”并回车\n然后将[signed.applets.codebase_principal_support]设置为'true'");}
				var prefs = Components.classes['@mozilla.org/preferences-service;1'].getService(Components.interfaces.nsIPrefBranch);
				prefs.setCharPref('browser.startup.homepage',url);
			}
		}
	},
	//打印 必须参数：sprnstr开始标记 eprnstr结束标记
	doPrint:function() 
	{    
		var sprnstr=arguments[0];
		var eprnstr=arguments[1];
		bdhtml=window.document.body.innerHTML;     
		prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+sprnstr.length);    
		prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr));    
		window.document.body.innerHTML=prnhtml;    
		window.print();    
	},  
	//添加收藏
	myAddBookmark:function() 
	{ 
		var title=arguments[0];
		var url=arguments[1];
		if ((typeof window.sidebar == 'object') && (typeof window.sidebar.addPanel == 'function'))//Gecko 
		{ 
			window.sidebar.addPanel(title,url,""); 
		} 
		else//IE 
		{ 
			window.external.AddFavorite(url,title); 
		} 
		return false;
	},
	//功能：添加下拉列表选项 _obj 显示分类的控件,_value1选项的显示值,_value2选项的value值
	addClass:function(){
		var _obj=arguments[0];
		var _value1=arguments[1];
		var _value2=arguments[2];
		var obj=eval("document."+_obj);
		//alert(typeof(obj));
		obj.options[obj.options.length]=new Option(_value1,_value2);			
	},
	//功能：选择下拉列表中的一个选项 
	selectClass:function()
	{
		var _obj=arguments[0];
		var _value2=arguments[1];
		var obj=eval("document."+_obj);
		for (var i = 0; i < obj.options.length; i++) {        
			if (obj.options[i].value == _value2) {        
				obj.options[i].selected = true;  
				break;        
			}  
		}
	},

	//
	chkfocusValue:function()
	{	
		var _obj=arguments[0];
		if(_obj.value==_obj.defaultValue) {_obj.value='';}
	},

	chkblurValue:function()
	{	
		var _obj=arguments[0];
		if(_obj.value=='') {_obj.value=_obj.defaultValue;}
	},

	//图片等比缩放
	pic_reset:function() {    
		var drawImage=arguments[0];
		var thumbs_size=arguments[1];
		var max = thumbs_size.split(',');    
		var fixwidth = max[0];    
		var fixheight = max[1];  
		var w=drawImage.width;
		var h=drawImage.height;    
		if(w>fixwidth) { drawImage.width=fixwidth;drawImage.height=h/(w/fixwidth);}    
		if(h>fixheight) { drawImage.height=fixheight;drawImage.width=w/(h/fixheight);}          
	   // drawImage.style.cursor= "pointer";    
	//    drawImage.onclick = function() { window.open(this.src);}    
	//    drawImage.title = "点击查看原始图片";  
	},

	//选中/取消选中 指定ID的复选按钮
	//可选参数 mode 0反选 1全选/全不选
	CheckAll:function()  {
		var _id=arguments[0];
		var _mode=0;
		var _obj;
		if(arguments.length==3){
			_mode=parseInt(arguments[1]);
			_obj=arguments[2];
		}
		var objArr=document.all(_id)
	  	for (var i=0;i<objArr.length;i++)    {
			var e = objArr[i]; 
			switch(_mode){
				case 0:{e.checked=!e.checked;break;}
				case 1:{e.checked=_obj.checked;break;}			
				//default:{e.checked =!e.checked;break;}
			}
	   	}
	},
	
  	//检查复选框是否至少有一个选择 返回true false 指定ID的复选按钮
	chkSelect:function()  {
		var f=false;
		var _id=arguments[0];
		var objArr=document.all(_id);
		if(objArr==null){return false;}		
		if(typeof(objArr.length)=="undefined"){
			if(objArr.checked) f=true;
		}else{
			for (var i=0;i<objArr.length;i++)    {
				var e = objArr[i];         
				if(e.checked){f=true;break;} 
			}
		}
		return f;
	},
	//显示提示信息 如果只有一个参数，只用alert显示，如果有两个参数，则在指定的ID(第一个参数为ID，第二个参数为消息)上显示
	showmsg:function()
	{
		if(arguments.length==1){
			alert(arguments[0]);
		}else{
			document.getElementById(arguments[0]).innerHTML=arguments[1];
		}
	},

	//检查cookiesname是否有值 _min最小长度 _max最大长度(为0时无限最大长度)
	//返回true 已经登录 false没有登录
	chkCookiesValue:function()
	{
		var cookiesname=arguments[0];
		var _min=arguments[1];
		var _max=arguments[2];
		var v=this.GetCookies(cookiesname);
		if(v!=''&&v.length>=_min)
		{
			if(_max==0)
				return true;
			else
			{
				if(v.length<=_max)
					return true;
				else
					return false;
			}
		}
		else
			return false;	
	},
	//获取cookies值
	GetCookies:function() 
	{ 
		var cookiesName=arguments[0];
		var varName ;
		varName= cookiesName; 
		if(varName==""){
			return "?";
		}
		try{
		//取 cookie 字符串，由于 expires 不可读，所以 expires 将不会出现在 cookieStr 中。
		var cookieStr = document.cookie; 
		}catch(e){
			alert("failue");
			}
		//alert(cookieStr);
		if (cookieStr == ""){
			//没有取到 cookie 字符串，返回默认值 
			return "?";
		} 
		//将各个 cookie 分隔开，并存为数组，多个 cookie 之间用分号加空隔隔开，不过前面我们只使用了一个 cookie，它的值与 expires 之间也是用分号加空格隔开	
		var cookieValue = cookieStr.split("; ");  
		
		for (var i=0; i<cookieValue.length; i++){
			if(cookieValue[i].indexOf("&")==-1){
				var cookiesName2=cookieValue[i].split("=");	
				if(cookiesName2.length>1){
					if (varName == cookiesName2[cookiesName2.length-2].toString())
						return unescape(cookiesName2[cookiesName2.length-1]);			
				}
			}else{
				var cookiesName2=cookieValue[i].split("&");	
				for(var j=0;j<cookiesName2.length;j++){
					var cookiesName3=cookiesName2[j].split("=");
					if(cookiesName3.length>1){
						if (varName == cookiesName3[cookiesName3.length-2].toString()) 
							return unescape(cookiesName3[cookiesName3.length-1]);					
					}
				}
			}		
				
		} 
		return "?"; 
	},

	//显示浮动窗口提示(调用框架) 
	//v 1为显示 其它值为关闭 title 浮动窗口标题 url框架的网址
	boxs:function(v,title,url){
		var v=arguments[0];
		var title=arguments[1];
		var url=arguments[2];
		document.charset   =   "utf-8";
		var ht = document.getElementsByTagName('html')[0];  	
		if(ht==null){
			ht=document.createElement("html");
			document.appendChild(ht);		
		}
		var bo = document.getElementsByTagName('body')[0];
		if(bo==null){
			bo=document.createElement("body");
			document.appendChild(bo);		
		}	
		var boht=document.getElementById("msgBox");	
		if(boht==null){		
			boht = document.createElement("span");
			boht.id="msgBox";
			bo.appendChild(boht); 		
		}
		
		window.scrollTo(0,0);	
		boht.innerHTML = '';
		bo.style.height='auto';
		bo.style.overflow='auto';
		ht.style.height='auto'; 
		if(v == 1){   
		bo.style.height='100%';
		bo.style.overflow='hidden';
		bo.style.margin='0';
		ht.style.height='100%'; 	
		boht.innerHTML = '<div style="background:#333333; opacity: 0.5;-moz-opacity:0.5; filter:alpha(opacity=50); width:100%; height:100%;position:absolute; top:0;left:0;z-index:99"></div><div style="z-index:999;height:0px; width:0px;top:50%; left:50%;position:absolute; line-height:1.7"><div style="background:#fff; border:1px solid #357bb9; width:400px; height:250px; position:absolute; margin:-150px -200px"><strong style="display:block; padding:2px 5px; background:#eef8f9; color:#4d4d4d; font-weight:100; width:390px; height:22px"><span style="float:left">'+title+'</span><span style="float:right"><a href="javascript:boxs(0);">关闭</a></span></strong><p style="padding:10px"><iframe frameborder="0" scrolling="no" width="380" height="200px" src="'+url+'"></iframe></p></div></div>';   	
		} 
	},
	//设置direct cookies
	setDirect:function()
	{
		var Days = 1;   
		var exp1 = new Date();    
		exp1.setTime(exp1.getTime() + Days*24*60*60*1000);  
		alert(document.domain);
		var h=window.href.toLowerCase();
		var h1=h.split("/");
		document.cookie="direct=/"+h1[h1.length-1]+";path=/;expires=" + exp1.toGMTString()+";";
	},
	
	//发送表单数据 formID:表单ID action:处理页面 resultID返回结果控件ID href操作成功返回的页面 为空不返回
	//本函数需要 jquery.js jquery.form.js两个库函数
	sendData:function() {  
		var formID=arguments[0];
		var action=arguments[1];
		var resultID=arguments[2];
		var href=arguments[3];
		var btobj=null;
		if(arguments.length==5){btobj=arguments[4]}
		var datas = $("#"+formID).formToArray();        //用了jquery.form.js 插件里面的一个方法,把表单数据拼接起		
		$.ajax({ 
			type: "POST", 
			url: action, 
			data: datas,
			beforeSend:function(XMLHttpRequest){
				$("#"+resultID).css("display","");
				btobj.disabled=true;
			},
			success: function(msg) { 
			   if(msg.indexOf("成功")!=-1||msg.indexOf("suc")!=-1){	
			   if(href!='') location.href=href;	
			   $("#"+resultID).html(msg);
			   }else{$("#"+resultID).html("产生错误，可能原因："+msg);btobj.disabled=false;}
			},
			complete: function(XMLHttpRequest, textStatus){
			//HideLoading();
			},
			error:function(){
				$("#"+resultID).html("产生错误！脚本页面有错误！");
				btobj.disabled=false;
			}
		}); 
		return false;             
	}
}