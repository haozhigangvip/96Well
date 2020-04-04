package com.hzg.servlet;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.http.Cookie;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import net.sf.json.JSONObject;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;

import com.hzg.domain.inExcel;
import com.hzg.services.ExcelServices;
import com.hzg.utils.ExcelUtils;
import com.hzg.utils.getuuid;


/**
 * Servlet implementation class ReadExcelServlet
 */
public class getNewExcelServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		//寰楀埌涓婁紶鏂囦欢鐨勪繚瀛樼洰褰曪紝灏嗕笂浼犵殑鏂囦欢瀛樻斁浜嶹EB-INF鐩綍涓嬶紝涓嶅厑璁稿鐣岀洿鎺ヨ闂紝淇濊瘉涓婁紶鏂囦欢鐨勫畨鍏�
		String savePath = this.getServletContext().getRealPath("/WEB-INF/upload");
		DiskFileItemFactory factory = new DiskFileItemFactory();
		String message="";
		//2銆佸垱寤轰竴涓枃浠朵笂浼犺В鏋愬櫒
		ServletFileUpload upload = new ServletFileUpload(factory);
		//瑙ｅ喅涓婁紶鏂囦欢鍚嶇殑涓枃涔辩爜
		upload.setHeaderEncoding("UTF-8"); 
		ExcelServices Excel=new ExcelServices() ;
		
		List<FileItem> list=null;
		try {
			list = upload.parseRequest(request);
	
			inExcel xls=Excel.readExcel(list,savePath);		
			
			//writer.flush();
			if(xls==null||xls.isReadError()){
				message="status: 0";


			}else
			{
				
				HttpSession session=request.getSession();
				
				String cookstr="{\"prows\":"+xls.getRows()+
						",\"pcols\":"+xls.getCols()+
						",\"margin_left\":"+xls.getMargin_left()+
						",\"margin_right\":"+xls.getMargin_right()+
						",\"margin_top\":"+xls.getMargin_top()+
						",\"margin_butto\":"+xls.getMargin_butto()+
						"}";
				
				Cookie ck=new Cookie("96wellCookie", cookstr);
				ck.setMaxAge(31104000);
				ck.setPath(request.getContextPath()+"/");
				response.addCookie(ck);
				
				
				String newfile=Excel.toExcel(request,xls);
				ServletContext con=this.getServletContext();
				message="status: 1,url:\""+request.getContextPath()+"/download/"+newfile+"\"";
				//message="true";
//				Excel.download(newfile, response, request, con);
				
			}
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			
		}finally {
			PrintWriter writer = response.getWriter();
			writer.write("{");
//			writer.write("msg:\"文件大小:"+item.getSize()+",文件名:"+filename+"\"");
//			writer.write(",picUrl:\"" + picUrl + "\"");
			writer.write(message);
			writer.write("}");
			writer.flush();
			
			writer.close();

		}
		
		
		
		
		
	
			
				
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request,response);
		
	}
	
	private String getUploadFileName(FileItem item) {
		// 获取路径名
		String value = item.getName();
		// 索引到最后一个反斜杠
		int start = value.lastIndexOf("/");
		// 截取 上传文件的 字符串名字，加1是 去掉反斜杠，
		String filename = value.substring(start + 1);
		
		return filename;
	}
	
	

}

