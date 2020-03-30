/**
 * 
 */
	function dispAlert(Content,curl){			
				$("#content").html(Content);	
				document.getElementById('zhezhao').style.display="block";
				$("#contentHref").on("click",function(){
					if(curl!=null && curl!=""){
					var nurl=curl;
					window.location.href=nurl;
					hiddAlert();}
				});	
			}
			
			function hiddAlert(){
				document.getElementById('zhezhao').style.display="none";
				document.getElementById('mainform').style.display="";
			}
			
