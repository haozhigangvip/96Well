package cn.hzg.Service;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import javax.servlet.ServletContext;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;
import cn.hzg.pojo.DataInfo;
import cn.hzg.pojo.plate;
import cn.hzg.Utils.ExcelUtils;
import cn.hzg.Utils.getuuid;;

public class ExcelServices {

	public List<plate> readExcel(MultipartFile file,String savePath) {
		// TODO Auto-generated method stub
		
		List<plate> rds=null;
		File lfile = new File(savePath);
		if (!lfile.exists() && !lfile.isDirectory()) {
			
			lfile.mkdir();
		}

		
		InputStream in=null;
		FileOutputStream out=null;
		try{
			String filename=file.getOriginalFilename();
			String lfilename=filename.substring(filename.lastIndexOf("."));
			in= file.getInputStream();
			String newfile=getuuid.getUUID();
			out= new FileOutputStream(savePath + "\\" +newfile+lfilename );
			byte buffer[] = new byte[1024];
			int len = 0;
			while((len=in.read(buffer))>0){
				out.write(buffer, 0, len);
			}
			String ff=savePath+"\\"+newfile+lfilename;					
			out.flush();
			rds=new ExcelUtils().excelToList(ff);
			//开启多线程，删除文件；
			Thread thread = new FileDelete(ff);
			thread.start();	
			return rds;
			}
			catch (Exception e) {
				e.printStackTrace();
				return null;
			}
			finally {
				try {
					in.close();
					out.close();
				} catch (IOException e) {
					// TODO 自动生成的 catch 块
					e.printStackTrace();
				}
		
			}
		
	}

	@SuppressWarnings("deprecation")
	public String toExcel(HttpServletRequest request,DataInfo df) {
		
		// TODO Auto-generated method stub
		XSSFWorkbook book	=new ExcelUtils().getXLSXBook(request.getRealPath("/download")+"/template.xlsx");
		String rowxl="abcdefghijklmnopqrstuvwxyz";
		XSSFSheet sheet=book.getSheetAt(0);
		XSSFRow trow=null;
		XSSFCell tcell=null;
		int lmar=df.getMargin_left();
		int rmar=df.getMargin_right();
		int tmar=df.getMargin_top();
		int bmar=df.getMargin_butto();
		int cols=df.getCols();
		int rows=df.getRows();
		int nrow=0;
		int btrow=11;//标题行
		int rounds=0;
		int topjjrow=1;//上行行间距行
		int bottojjrow=1;//下行间距行
		int jrr=topjjrow+2;
		XSSFFont fonta =book.createFont();
		XSSFFont fontb =book.createFont();
		XSSFFont fontd =book.createFont();
		XSSFFont fonte =book.createFont();
		XSSFFont fontf =book.createFont();
		XSSFCellStyle empty_cs=(sheet.getRow(11).getCell(0).getCellStyle());
		XSSFCellStyle data_cs=(sheet.getRow(11).getCell(1).getCellStyle());
		Boolean nomarg=false;
		List<plate> list=df.getList();
		
		int zzrow=list.size();
		int listn=0;
		if(zzrow % ((cols-lmar-rmar)*(rows-tmar-bmar))==0){
			rounds=zzrow/((cols-lmar-rmar)*(rows-tmar-bmar));
		}else
		{
			rounds=(zzrow/((cols-lmar-rmar)*(rows-tmar-bmar)))+1;
		}
		
		XSSFCellStyle bqStyle=book.createCellStyle();
		XSSFCellStyle bq1Style=book.createCellStyle();
		XSSFCellStyle btaStyle=book.createCellStyle();
		XSSFCellStyle btbStyle=book.createCellStyle();
		XSSFCellStyle sjaStyle=(XSSFCellStyle)data_cs.clone();
		XSSFCellStyle sjbStyle=book.createCellStyle();
		XSSFCellStyle sjcStyle=(XSSFCellStyle)data_cs.clone();
		XSSFCellStyle sjdStyle=book.createCellStyle();
		XSSFCellStyle sjeStyle=(XSSFCellStyle)data_cs.clone();
		XSSFCellStyle sjfStyle=book.createCellStyle();
		XSSFCellStyle sjgStyle=(XSSFCellStyle)data_cs.clone();
		XSSFCellStyle sjhStyle=book.createCellStyle();
		XSSFCellStyle rowbtStyle=book.createCellStyle();
		XSSFCellStyle emptyaStyle=(XSSFCellStyle)empty_cs.clone();
		XSSFCellStyle emptycStyle=(XSSFCellStyle)empty_cs.clone();
		XSSFCellStyle emptybStyle=(XSSFCellStyle)empty_cs.clone();
		XSSFCellStyle emptydStyle=(XSSFCellStyle)empty_cs.clone();
		XSSFCellStyle emptyeStyle=(XSSFCellStyle)empty_cs.clone();
		

		for (int rnd=0;rnd<rounds;rnd++){
			for(int rr=0;rr<(rows+1)*2+1;rr++){
				
				nrow=btrow+rnd*(rows*2+jrr+bottojjrow)+rr;
		
				trow=sheet.createRow(nrow);
				
				if((rr-jrr)%2==0){
				trow.setHeight((short) (24*20));
				}else{
					trow.setHeight((short) (27.75*20));
				}

				for(int cc=0;cc<cols+1;cc++){
					nomarg=false;
					tcell=trow.createCell(cc);
					if(rr==0){	
						//设置Plate layout
						if(topjjrow>0){
							sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonta,bqStyle,0,0 ,1,12));
						}else
						{
							
							sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonta,bqStyle,2,0 ,1,12));
						}
						
						
						tcell.setCellValue("Plate layout:"+list.get(listn).getPlate());
						trow.setHeight((short) (14.3*20));
						}
					if(rr==topjjrow && topjjrow>0){
						sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonta,bq1Style,2,0 ,1,12));
					}
					if(rr==topjjrow+1){	
						//设置列标题行高
						trow.setHeight((short) (19.5*20));
						if(cc>0){
							//设置列标题
							tcell.setCellValue(cc);
							
						}
						
						//设置列标题边距及格式
						if(cc<cols){
						 //前几列
						sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontb, btaStyle,0,2 ,1,10));
						}
						else
						{
						 //最后列
							sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontb, btbStyle,4,2 ,1,10));
						}
					}
					
					if(rr>topjjrow+1){
						//设置行标题	
						if(cc==0){
							if((rr-jrr)%2==0){
							
							sheet.addMergedRegion(new CellRangeAddress(nrow,nrow+1,0,0));
							tcell.setCellValue(rowxl.substring((rr-jrr)/2, (rr-jrr)/2+1));
							}
							sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontb, rowbtStyle,4,2 ,1,10));
						 
						}else{
							
								if((rr-jrr)%2>0 && (rr-jrr>=tmar*2+1) && rr-jrr<(rows)*2-bmar*2){
									if(cc>lmar && cc<cols-rmar+1 && listn<list.size()){
										//填充CAS
										sheet.getRow(nrow-1).getCell(cc).setCellValue(list.get(listn).getCAS()); 
										
										//填充Compound
										sheet.getRow(nrow).getCell(cc).setCellValue(list.get(listn).getCompound());
										
										listn++;
									}
								if(cc<cols-rmar||rmar==0){
								//CAS格式
									if(rr-jrr==(tmar)*2+1 && tmar>0){
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonte,sjeStyle,10,2 ,1,8));
									}else
									{
										
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonte,sjaStyle,6,2 ,1,8));
									}	
								
								//Compund格式
								  if(rr-jrr==(rows)*2-bmar*2-1 && bmar>0){
									  sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontf,sjfStyle,11,2 ,0,7));
								  }else
								  {
								   sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontf,sjbStyle,7,2 ,0,7));
								  }
								}
								else
								{
								
								
								    //CAS格式
									if(rr-jrr==(tmar)*2+1 && tmar>0){
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonte,sjgStyle,14,2 ,1,8));
									}else{
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonte,sjcStyle,1,2 ,1,8));
									}
								
									
									//Compund格式
									 if(rr-jrr==(rows)*2-bmar*2-1 && bmar>0){
										 sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontf,sjhStyle,15,2 ,0,7));
									 }else
									 {
										 sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontf,sjdStyle,2,2 ,0,7));
									 }
									
								}
								}
								
							//设置列边距
							if(cc<=lmar ||cc>=cols-rmar+1){
								//设置边距列宽
								sheet.setColumnWidth(cc, (int)8.38*252+323);
								tcell.setCellValue("Empty");
								
								
								 if((rr-jrr)%2>0 ){
								 
									 sheet.addMergedRegion(new CellRangeAddress(nrow-1,nrow,cc,cc));
									 //设置列边距格式
									 if((cc==lmar  )){
										
										 sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptyaStyle,8,2 ,0,8));
										 sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptyaStyle,8,2 ,0,8));
									 }else if(cc==cols-rmar+1){
										
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptybStyle,9,2 , 0,8));
										sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptybStyle,9,2 ,0,8));
									 }else
									 {
									
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptycStyle,5,2 ,0,8));
										sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptycStyle,5,2 ,0,8));
									 }
									 
									 
									 nomarg=true;
								
								 }
								
							}else{
								//设置数据列宽
								sheet.setColumnWidth(cc, 11*252+323);
							}
							
							
							//设置行边距
							if(rr-jrr<tmar*2||rr-jrr>=(rows)*2-bmar*2){
								 tcell.setCellValue("Empty");
								 if((rr-jrr)%2>0){
									 if(nomarg==false){
									 sheet.addMergedRegion(new CellRangeAddress(nrow-1,nrow,cc,cc));
									 }
									 
									 //设置行边距格式
									  if(rr-jrr<=tmar*2-1||rr-jrr>=(rows)*2-bmar*2+1){
									  if(rr-jrr==(rows)*2-bmar*2+1 && bmar>0){
									   sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptydStyle,13,2 ,0,8));
									  }else
									  {
								      sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptycStyle,5,2 , 0,8));
									  }
								      if(rr-jrr==tmar*2-1 && tmar>0){
										  sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptyeStyle,12,2 ,0,8));
									  }else
									  { 
										  sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptycStyle,5,2 ,0,8));
									  }
								    
									 
										 
									  }
									 
								 }
							 }
						}
					}	
				}
				if(rr==0){
					sheet.addMergedRegion(new CellRangeAddress(nrow,nrow,0,cols));
				}
				if(topjjrow>0 && rr>=topjjrow &rr<=topjjrow){
					sheet.addMergedRegion(new CellRangeAddress(nrow,nrow,0,cols));
					
					trow.setHeight((short) (12.75*20));
				}
				
					
			}
			
		}
		
		
		//保存数据
		
		String filename=String.valueOf(System.currentTimeMillis())+".xlsx";
		FileOutputStream out=null;
		try {
			out = new FileOutputStream(request.getRealPath("/")+"download/"+filename);
			book.write(out);
			out.flush();
			out.close();
			return filename;
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return null;
		}finally{
			try {
				out.flush();
				out.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		
	
	}
	
	public void download(String file,HttpServletResponse response,HttpServletRequest request,ServletContext con) throws IOException{

	    //閽堝涓嶅悓娴忚鍣ㄦ敼鍙樼紪鐮�
	     
			//鑾峰彇conntext
			
			//璁剧疆鏂囦欢mimeType
			
			String mimetype=con.getMimeType(file);
			response.setContentType(mimetype);
			//璁剧疆涓嬭浇澶翠俊鎭�
			response.setHeader("content-disposition", "attchment;filename="+file);
			//瀵规嫹娴�
			//鑾峰彇杈撳叆娴�
			
			InputStream is=con.getResourceAsStream("/download/"+file);
			
			//鑾峰彇杈撳嚭娴�
			ServletOutputStream os=response.getOutputStream();
			
			int len=-1;
			byte[] b=new byte[1024];

			while((len=is.read(b))!=-1) {
				os.write(b,0,len);
			}
			
			os.flush();
			os.close();
			is.close();
			
	}
	

}
