document.write("<div class='QQbox' id='divQQbox' >");

document.write("<div class='Qlist' id='divOnline' onmouseout='hideMsgBox(event);' style='display : none;'>");

document.write("<div class='t'></div>");

document.write("<div class='con'>");

document.write("<h2>在线客服</h2>");

document.write("<ul>");

document.write("<li><a href=' http://wpa.qq.com/msgrd?V=1&amp;Uin=276033090&amp;Site=广安娱乐&amp;Menu=yes' target='_blank'><img src=' http://wpa.qq.com/pa?p=1:276033090:10'  border='0' alt='广安娱乐QQ在线交谈' /></a></li>");

document.write("<li><a href=' http://wpa.qq.com/msgrd?V=1&amp;Uin=841492922&amp;Site=广安娱乐&amp;Menu=yes' target='_blank'><img src=' http://wpa.qq.com/pa?p=1:841492922:10'  border='0' alt='广安娱乐QQ在线交谈' /></a></li>");

document.write('<li><a target="_blank" href="http://amos.im.alisoft.com/msg.aw?v=2&uid=%E5%8D%A1%E6%9F%A0%E5%8D%A1%E6%AA%AC&site=cntaobao&s=1&charset=utf-8" ><img border="0" src="http://amos.im.alisoft.com/online.aw?v=2&uid=%E5%8D%A1%E6%9F%A0%E5%8D%A1%E6%AA%AC&site=cntaobao&s=1&charset=utf-8" alt="阿里旺旺在线交谈" /></a></li>');
document.write('<li><a  href="msnim:chat?contact=gaylkj668@msn.com" >MSN立即交谈</a></li>');


document.write('<li style="color:red;font-size:12px;">Mobile:13660170912</li>');

document.write("</ul>");document.write("</div>");

document.write("<div class='b'></div>");

document.write("</div>");

document.write("<div id='divMenu' onmouseover='OnlineOver();'><img src='/skin/default/QQkefu/images/qq_1.png' class='press' alt='QQ客服热线'></div>");

document.write("</div>");



//<![CDATA[

var tips; var theTop = 145/*这是默认高度,越大越往下*/; var old = theTop;

function initFloatTips() {

tips = document.getElementById('divQQbox');

moveTips();

};

function moveTips() {

var tt=50;

if (window.innerHeight) {

pos = window.pageYOffset

}

else if (document.documentElement && document.documentElement.scrollTop) {

pos = document.documentElement.scrollTop

}

else if (document.body) {

pos = document.body.scrollTop;

}

pos=pos-tips.offsetTop+theTop;

pos=tips.offsetTop+pos/10;



if (pos < theTop) pos = theTop;

if (pos != old) {

tips.style.top = pos+"px";

tt=10;

//alert(tips.style.top);

}



old = pos;

setTimeout(moveTips,tt);

}

//!]]>

initFloatTips();







function OnlineOver(){

document.getElementById("divMenu").style.display = "none";

document.getElementById("divOnline").style.display = "block";

document.getElementById("divQQbox").style.width = "145px";

}



function OnlineOut(){

document.getElementById("divMenu").style.display = "block";

document.getElementById("divOnline").style.display = "none";



}


if(typeof(HTMLElement)!="undefined")    //给firefox定义contains()方法，ie下不起作用
{   
      HTMLElement.prototype.contains=function(obj)   
      {   
          while(obj!=null&&typeof(obj.tagName)!="undefind"){ //通过循环对比来判断是不是obj的父元素
   　　　　if(obj==this) return true;   
   　　　　obj=obj.parentNode;
   　　}   
          return false;   
      };   
}  


function hideMsgBox(theEvent){ //theEvent用来传入事件，Firefox的方式

　 if (theEvent){

　 var browser=navigator.userAgent; //取得浏览器属性

　 if (browser.indexOf("Firefox")>0){ //如果是Firefox

　　 if (document.getElementById('divOnline').contains(theEvent.relatedTarget)) { //如果是子元素

　　 return; //结束函式

} 

} 

if (browser.indexOf("MSIE")>0){ //如果是IE

if (document.getElementById('divOnline').contains(event.toElement)) { //如果是子元素

return; //结束函式

}

}

}

/*要执行的操作*/

document.getElementById("divMenu").style.display = "block";

document.getElementById("divOnline").style.display = "none";

}