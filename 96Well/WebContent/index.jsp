<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<%@taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link rel="stylesheet" href="${pageContext.request.contextPath }/css/bootstrap.min.css">
<link rel="stylesheet" href="${pageContext.request.contextPath }/css/style.default.css" id="theme-stylesheet">
<link rel="stylesheet" href="${pageContext.request.contextPath }/css/PopupWindow.css" >
<script type="text/javascript" src="${pageContext.request.contextPath }/js/jquery.js"></script>
<script type="text/javascript" src="${pageContext.request.contextPath }/js/ajaxfileupload.js"></script>
<script type="text/javascript" src="${pageContext.request.contextPath }/js/PopupWindow.js"></script>
<script type="text/javascript" src="${pageContext.request.contextPath }/js/jquery.cookie.js"></script>

<title>96-WELL</title>
</head>

<body>

		
		<div class="zhezhao" id='zhezhao'>
			<div class="tankuang">
				<div id="header" >
					<span>系统提示</span>
					<div id="header-right" onclick="hiddAlert();">x</div>
				</div>
				<div class="card-body" align="center">
				<table>
				<tr>
				<td>
				<div class="input-group" style="height: 50px"><label id="content"><input type="text" id="alterurl"></label></div>
				</td>
				</tr>				
				<tr >				
				<td align="center">
				<button type="button" onclick="hiddAlert();">关    闭</button>
				</td>				
				</tr>				
				</table>
			  </div>	
			</div>
		</div>


        <div class="container-fluid px-xl-5" style="margin-top:40px;margin-left:300px" id="mainform">

         
              <!-- Basic Form-->
              <div class="col-lg-6 mb-5">
                <div class="card">
                  <div class="card-header">
                    <h3 class="h6 text-uppercase mb-0">Bioactive Compound Library  (96-well)</h3>
                  </div>
                  <div class="card-body">
                   <form method="post" action="" enctype="multipart/form-data" id="myForm" name="myForm">
                   
                    <table>
 					 <tr height="60px">
						<td colspan="2">
	        				    <div class="input-group">
								<label class="form-control-label text-uppercase">选择文件</label>
								<input class="form-control" id="upfile" type="file"  name="file" accept=".xls,.xlsx,.xlsm" width="300px" onchange="Check();">
								</div>
						</td>
 					 </tr>
 					 <tr>
 				  		<td colspan="2">
 				  			<div class="invalid-feedback" id="fileErr" >请选择Excel文件</div>
 				  		</td>
 				  	</tr>
 					 
 					<tr height="60px">
						<td>
						<div class="input-group">
								
					  	 	 <label class="form-control-label text-uppercase"  >行数</label>
							
							<input type="text" placeholder="输入行数" class="form-control"  name="prows"   id="prows" onkeyup="Check();"  > 	
						  </div>
	
						
						</td>
						<td>
							<div class="input-group">
							 <label class="form-control-label text-uppercase">列数</label> 	
                      	 	 <input type="text" placeholder="输入列数" class="form-control"  name="pcols"   id="pcols" onkeyup="Check();">
                      	 	 </div>
						</td>
 				  </tr>
 				  <tr>
 				  		<td>
 				  			<div class="invalid-feedback" id="prowsErr" >请输入行数</div>
 				  		</td>
 				  		<td>
 				  		<div class="invalid-feedback" id="pcolsErr" >请输入列数</div>
 				  		</td>
 				  </tr>
 				  
 				  
				  <tr height="60px">
					<td>
						<div class="input-group">
					  	 	 <label class="form-control-label text-uppercase">左边距</label>
                   	    	  <input type="text" placeholder="输入左边距" class="form-control" value=0 name="margin_left" id="margin_left" onkeyup="Check();">
                   	    </div>
					</td>
					<td>
						<div class="input-group">
							<label class="form-control-label text-uppercase">右边距</label>
                      	  	<input type="text" placeholder="输入右边距" class="form-control" value=0 name="margin_right"  id="margin_right" onkeyup="Check();" >
                      	  </div>
					</td>
				  </tr>
				  <tr>
 				  		<td colspan="2">
 				  			<div class="invalid-feedback" id="marginleftrightErr" >左右边距之和必须小于列数</div>
 				  		</td>
 				  		
 				  </tr>
				  
				  
				  
				  
				  
				  <tr height="60px">
					<td>
						<div class="input-group">
					  	 	 <label class="form-control-label text-uppercase">上边距</label>
                   	    	  <input type="text" placeholder="输入上边距" class="form-control" value=0 name="margin_top"  id="margin_top" onkeyup="Check();">
                   	    </div>
					</td>
					<td>
						<div class="input-group">
							<label class="form-control-label text-uppercase">下边距</label>
                      	 	 <input type="text" placeholder="输入下边距" class="form-control" value=0 name="margin_butto"  id="margin_butto" onkeyup="Check();">
						  </div>

					</td>
				  </tr>
				   <tr>
 				  		<td colspan="2">
 				  			<div class="invalid-feedback" id="margintopbuttoErr" >上下边距之和必须小于行高</div>
 				  		</td>
 				  		
 				  </tr>
				  <tr height="60px">
					<td colspan="2">
						<div class="input-group">  
						     
                       		 <button type="submit" class="btn btn-primary" style="margin:auto 20px" id="sbtn" onclick="return ajaxFileUpload();">转换</button>
					   
					   		<div id="loading" class="loading" style="display:none" >
					   			<div class="loadingimg">
					   		
								<p><img src="img/timg.gif" />
								<p><a>正在转换,请稍后...</a>
                   				</div>
                   			</div>
                   		</div>
                   </td>
                      
				  </tr>
				  
				 
				  
				  
 			 </table>
 			
		 </form>

                  </div>
                </div>
		
   		 </div>
	</div>
					 
</body>

<script type="text/javascript">
function SetErrorClass(id,ErrDiv_Id){
	$(id).attr("class","form-control is-invalid");

	$(ErrDiv_Id).show();
	
	
}
function SetTrueClass(id,ErrDiv_Id){
	$(id).attr("class","form-control");

	$(ErrDiv_Id).hide();
	
	
}
function Check(){
	 var file= $("#upfile").val();
	 var prows=Number($("#prows").val());
	 var pcols=Number($("#pcols").val());
	 var margin_right=Number($("#margin_right").val());
	 var margin_left=Number($("#margin_left").val());
	 var margin_top=Number($("#margin_top").val());
	 var margin_butto=Number($("#margin_butto").val());
	 var rt=true;
	
	 if(file==null || file==""){
		 SetErrorClass("#upfile","#fileErr");
		 rt=false;
	 }else
	{
		 SetTrueClass("#upfile","#fileErr");
	}	 
	 
	
	 if(prows==0){
		
		SetErrorClass("#prows","#prowsErr");
		rt= false;
	}else{
		SetTrueClass("#prows","#prowsErr");
	}
		
	 
	 
	 if(pcols==0){
		SetErrorClass("#pcols","#pcolsErr");
		rt= false;
	}else{
		SetTrueClass("#pcols","#pcolsErr");
	}
	 
	 if(margin_left+margin_right>=pcols && margin_left+margin_right>0 ){
		 SetErrorClass("#margin_left","#marginleftrightErr");
		 SetErrorClass("#margin_right","#marginleftrightErr");
		rt= false;
	}else{
		 SetTrueClass("#margin_left","#marginleftrightErr");
		 SetTrueClass("#margin_right","#marginleftrightErr");
	}
	 
	 
	 
	 
	 if(margin_top+margin_butto>=prows && margin_top+margin_butto>0 ){
		 SetErrorClass("#margin_top","#margintopbuttoErr");
		 SetErrorClass("#margin_butto","#margintopbuttoErr");
		rt= false;
	}else
	{
		 SetTrueClass("#margin_top","#margintopbuttoErr");
		 SetTrueClass("#margin_butto","#margintopbuttoErr");
	}
	
	return rt;
}
function ajaxFileUpload(){
	 var file= $("#upfile")[0].files[0];
		
	 var rows=parseInt($("#prows").val());
	 var cols=parseInt($("#pcols").val());
	 var margin_right=parseInt($("#margin_right").val());
	 var margin_left=parseInt($("#margin_left").val());
	 var margin_top= parseInt($("#margin_top").val());
	 var margin_butto=parseInt($("#margin_butto").val());
	
	 var formData=new FormData();


	if(Check()==false){

		return false;
	}; 
	
	formData.append("file",file);

	var DataInfo= JSON.stringify({
	    "rows":rows,
	    "cols": cols,
	    "margin_right": margin_right,
	    "margin_left": margin_left,
	    "margin_top": margin_top,
	    "margin_butto":margin_butto
	    
	});

	formData.append('DataInfo', new Blob([DataInfo],{type: "application/json"}));
	
	$("#sbtn").blur();
	$("#loading").show();

	$.ajax({
	   			type: "post",	
				url: "getNewExcel.action",	
				processData: false,
				contentType : false,
				dataType : "json",
				data : formData,
				cache: false,
				success: function(data) {
					valid = false;
					if(data.status==0){
						$("#loading").hide(); 
						dispAlert("数据格式错误,请确认后重新转换","");
						
					}else
					{
						valid = false;
						var strs = new Array();
						var url=data.url;
						strs=url.split('/');
						var filename=strs[strs.length-1];
						dispAlert('转换完成，请点击'+"<a style=\"color:red\" href=\"javascript:void(0);\" id=\"contentHref\">"+filename+"</a>下载",url);
					}
					$("#loading").hide();  
				}, 
				error: function(e) {
					$("#loading").hide();
					valid = false;
					dispAlert("系统错误:"+ e +"，请联系管理员","");
				}
			})
			return false;	
}



</script>


<script type="text/javascript">
//读取cookie 设置
$(function(){

	if($.cookie("96wellCookie")!=null){
		
		ck1=decodeURI($.cookie("96wellCookie"));
		ck=JSON.parse(ck1);
		if(ck.prows!=null){document.myForm.prows.value=ck.prows;}
		if(ck.pcols!=null){document.myForm.pcols.value=ck.pcols;}
		if(ck.margin_left!=null){document.myForm.margin_left.value=ck.margin_left;}
		if(ck.margin_right!=null){document.myForm.margin_right.value=ck.margin_right;}
		if(ck.margin_top!=null){document.myForm.margin_top.value=ck.margin_top;}	
		if(ck.margin_butto!=null){document.myForm.margin_butto.value=ck.margin_butto;}
	}

		
});


</script>


</html>

