<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link rel="stylesheet" href="css/bootstrap.min.css">
<link rel="stylesheet" href="css/style.default.css" id="theme-stylesheet">
<link rel="stylesheet" href="css/PopupWindow.css" >
<script type="text/javascript" src="js/jquery.js"></script>
<script type="text/javascript" src="js/ajaxfileupload.js"></script>
<script type="text/javascript" src="js/PopupWindow.js"></script>

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
                   <form method="post" action="" enctype="multipart/form-data" id="myForm">
                   
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
					  	 	 <label class="form-control-label text-uppercase"  >行高</label>
                   	    	  <input type="text" placeholder="输入行高" class="form-control" name="prows" id="prows" onkeyup="Check();">
							
						  </div>
	
						
						</td>
						<td>
							<div class="input-group">
							 <label class="form-control-label text-uppercase">列宽</label> 	
                      	 	 <input type="text" placeholder="输入列宽" class="form-control"  name="pcols"   id="pcols" onkeyup="Check();">
                      	 	 </div>
						</td>
 				  </tr>
 				  <tr>
 				  		<td>
 				  			<div class="invalid-feedback" id="prowsErr" >请输入行高</div>
 				  		</td>
 				  		<td>
 				  		<div class="invalid-feedback" id="pcolsErr" >请输入列宽</div>
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
 				  			<div class="invalid-feedback" id="marginleftrightErr" >左右边距之和必须小于列宽</div>
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
	 
	 if(margin_left+margin_right>=pcols && margin_left+margin_right>0){
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
	 var file= $("#upfile").val();
	 var prows=Number($("#prows").val());
	 var pcols=Number($("#pcols").val());
	 var margin_right=Number($("#margin_right").val());
	 var margin_left=Number($("#margin_left").val());
	 var margin_top=Number($("#margin_top").val());
	 var margin_butto=Number($("#margin_butto").val());
	 
	if(Check()==false){

		return false;
	};
	
	$("#sbtn").blur();

	$("#loading").show();
	$.ajaxFileUpload({
	   		type: 'post',
				url: 'getNewExcel',
				secureuri : false,
				fileElementId : 'upfile',
				dataType : 'json',
				data: {prows : prows,
	       		pcols : pcols,
	       		margin_right : margin_right ,
	       		margin_left : margin_left,
	       		margin_top : margin_top,
	       		margin_butto :margin_butto},
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
				error: function(data,e) {
					$("#loading").hide();    
					valid = false;
					dispAlert("系统错误:"+ e +"，请联系管理员","");

				}
			})

			return false;	

}



</script>





</html>
