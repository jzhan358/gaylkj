// JavaScript Document
//获取浏览次数
function getViewNum(type,id)
{	
	$.ajax({
	type:"get",	
	dataType:"html",	
	cache:false,
	url:"/getViewNum.asp?type="+type+"&id="+id,
	beforeSend:function(XMLHttpRequest){
		$("#viewnum").html("<img src='/jquery/images/loading2.gif' border='0'/>");
		},
	success:function(data, textStatus){
		$("#viewnum").html(data);
		},
	complete: function(XMLHttpRequest, textStatus){
			//HideLoading();
		},
	error: function(){
			//alert("产品评论结果可能没有显示");
		}
	
	});
}