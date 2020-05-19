package cn.hzg.Controller;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import javax.servlet.http.Cookie;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import cn.hzg.Service.ExcelServices;
import cn.hzg.pojo.DataInfo;
import cn.hzg.pojo.plate;

@Controller
public class to96Excel {

	@RequestMapping("/getNewExcel")
	public  @ResponseBody String json(@RequestPart("DataInfo") DataInfo df,@RequestPart("file") MultipartFile file,HttpServletRequest request,HttpServletResponse response) throws UnsupportedEncodingException{
		String savePath = request.getServletContext().getRealPath("/WEB-INF/upload");
		String message="";
		List<plate> list=null;
		ExcelServices Excel=new ExcelServices();
		list=Excel.readExcel(file,savePath);	
		
		if(list==null){
			message="status: 0";
		}else
		{
			
			df.setList(list);
			String cookstr="{\"prows\":"+df.getRows()+
					",\"pcols\":"+df.getCols()+
					",\"margin_left\":"+df.getMargin_left()+
					",\"margin_right\":"+df.getMargin_right()+
					",\"margin_top\":"+df.getMargin_top()+
					",\"margin_butto\":"+df.getMargin_butto()+
					"}";
			String encodeCookie = URLEncoder.encode(cookstr,"UTF-8");
			Cookie ck=new Cookie("96wellCookie", encodeCookie);
			ck.setMaxAge(31104000);
			ck.setPath(request.getContextPath()+"/");
			System.out.println(request.getContextPath());
			
			response.addCookie(ck);
			String newfile=Excel.toExcel(request,df);
			message="\"status\": 1,\"url\":\""+request.getContextPath()+"/download/"+newfile+"\"";

		}
		message="{"	+ message +"}";
		return message;
	}
	
}
