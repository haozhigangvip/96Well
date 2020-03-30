<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Ajax File Uploader Plugin For Jquery</title>
<script type="text/javascript" src="jquery.js"></script>
<script type="text/javascript" src="ajaxupload.js"></script>
<script type="text/javascript">
$(function(){
	//上传图片
	new AjaxUpload('#addLabProdPic', {
		action: 'upload', 
		name: 'picFile',
		responseType: 'json',
		onSubmit : function(file , ext){
			if (ext && /^(jpg|png|bmp)$/.test(ext.toLowerCase())){
				this.setData({
					'picName': file
				});
			} else {
				alert("请上传格式为 jpg|png|bmp 的图片！");
				return false;				
			}
		},
		onComplete : function(file,response){
			if(response.error) {
				alert(response.error);
				return;
			}
			$('#viewImg').attr('src',response.picUrl);
		}		
	});
})
</script>
</head>

<body>
	<h1>Ajax文件上传例子,选择图片后立即上传,无需点击上传按钮</h1>
	<button id="addLabProdPic">选择图片</button>
	<br>
	<img id="viewImg">
</body>
</html>