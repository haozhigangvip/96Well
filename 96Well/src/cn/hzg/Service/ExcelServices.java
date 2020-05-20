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
			//寮�惎澶氱嚎绋嬶紝鍒犻櫎鏂囦欢锛�
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
					// TODO 鑷姩鐢熸垚鐨�catch 鍧�
					e.printStackTrace();
				}
		
			}
		
	}

	@SuppressWarnings("deprecation")
	public String toExcel(HttpServletRequest request,DataInfo df) {
		
		// TODO Auto-generated method stub
		XSSFWorkbook book	=new ExcelUtils().getXLSXBook(request.getRealPath("/WEB-INF/template")+"/template.xlsx");
		String rowxl="abcdefghijklmnopqrstuvwxyz";
		XSSFSheet sheet=book.getSheetAt(0);
		XSSFRow trow=null;
		XSSFCell tcell=null;
		int lmar=df.getMargin_left();
		int rmar=df.getMargin_right();
		int tmar=df.getMargin_top();
		int bmar=df.getMargin_butto();
		int mv=6;

		int mvv=(mv>0?1:0);
		
		int cols=df.getCols();
		int rows=df.getRows();
		int nrow=0;
		int btrow=11;//鏍囬琛�
		int rounds=0;
		int topjjrow=1;//涓婅琛岄棿璺濊
		int bottojjrow=1;//涓嬭闂磋窛琛�
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
		if(zzrow % ((cols-lmar-rmar-mvv)*(rows-tmar-bmar))==0){
			rounds=zzrow/((cols-lmar-rmar-mvv)*(rows-tmar-bmar));
		}else
		{
			rounds=(zzrow/((cols-lmar-rmar-mvv)*(rows-tmar-bmar)))+1;
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
		XSSFCellStyle emptyfStyle=(XSSFCellStyle)empty_cs.clone();
		XSSFCellStyle emptygStyle=(XSSFCellStyle)empty_cs.clone();

		for (int rnd=0;rnd<rounds;rnd++){
			for(int rr=0;rr<(rows+1)*2+1;rr++){
				
				nrow=btrow+rnd*(rows*2+jrr+bottojjrow)+rr;
				System.out.println("nrow:"+nrow);
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
						
						//填充Plate layout
						tcell.setCellValue("Plate layout:"+list.get(listn).getPlate());
						trow.setHeight((short) (14.3*20));
						}
					
					//设置标题顶端行格式
					if(rr==topjjrow && topjjrow>0){
						sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonta,bq1Style,2,0 ,1,12));
					}
					
					if(rr==topjjrow+1){	
						//填充列标题
						trow.setHeight((short) (19.5*20));
						if(cc>0){
							 tcell.setCellValue(cc);
						}
						//设置列标题格式
						if(cc<cols){
						 //列标题除最后一个格式
						sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontb, btaStyle,0,2 ,1,10));
						}
						else
						{
						 //列标题最后一个格式
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
							
							
								if((rr-jrr)%2>0 && (rr-jrr>=tmar*2+1) && rr-jrr<(rows)*2-bmar*2 ){
									if(cc>lmar && cc<cols-rmar+1 && cc!=mv && listn<list.size()){
										//填充CAS
										sheet.getRow(nrow-1).getCell(cc).setCellValue(list.get(listn).getCAS()); 
										
										//填充Compound
										sheet.getRow(nrow).getCell(cc).setCellValue(list.get(listn).getCompound());
										
										listn++;
									}
									
									
								if((cc<cols-rmar||rmar==0)&& cc!=mv){
								//设置CAS格式									
									if(rr-jrr==(tmar)*2+1 && tmar>0){
										//EMPYT下一行的首行
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonte,sjeStyle,10,2 ,1,8));
									}else
									{
										//非EMPYT下一行的首行
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonte,sjaStyle,6,2 ,1,8));
									}	
								
								//设置Compund格式
								  if(rr-jrr==(rows)*2-bmar*2-1 && bmar>0){
										//EMPYT下一行的首行
									  sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontf,sjfStyle,11,2 ,0,7));
								  }else
								  {
										//EMPYT下一行的首行
									  sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontf,sjbStyle,7,2 ,0,7));
								  }
								}
								else
								{
								
								
								    //设置最右侧首列CAS样式
									if(rr-jrr==(tmar)*2+1 && tmar>0){
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonte,sjgStyle,14,2 ,1,8));
									}else{
										sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonte,sjcStyle,1,2 ,1,8));
									}
								
									
									//设置最右侧首列Compund样式
									 if(rr-jrr==(rows)*2-bmar*2-1 && bmar>0){
										 sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontf,sjhStyle,15,2 ,0,7));
									 }else
									 {
										 sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontf,sjdStyle,2,2 ,0,7));
									 }
									
								}
								}
								
							//填充empty
							if(cc<=lmar ||cc>=cols-rmar+1|| cc==mv){
								//璁剧疆杈硅窛鍒楀
								sheet.setColumnWidth(cc, (int)8.38*252+323);
								tcell.setCellValue("Empty");
								
								
								 if((rr-jrr)%2>0 ){
								 
									 sheet.addMergedRegion(new CellRangeAddress(nrow-1,nrow,cc,cc));
									 
									
									//设置empty右虚线
									 if(cc==lmar  || (cc==mv && mv>0) ){
										if(cc==lmar){
										 sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptyaStyle,8,2 ,0,8));
										 sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptyaStyle,8,2 ,0,8));
										}
										if(cc==mv && mv>0 ){
											System.out.println("cc:"+cc+",nrow:"+nrow);
											 sheet.getRow(nrow-1).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptygStyle,16,2 ,0,8));
											 sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontd,emptygStyle,16,2 ,0,8));
										 }
									 }else if(cc==cols-rmar+1){
										//设置最右边格式
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
								//璁剧疆鏁版嵁鍒楀
								sheet.setColumnWidth(cc, 11*252+323);
							}
							
							
							//璁剧疆琛岃竟璺�
							if(rr-jrr<tmar*2||rr-jrr>=(rows)*2-bmar*2){
								 tcell.setCellValue("Empty");
								 if((rr-jrr)%2>0){
									 if(nomarg==false){
									 sheet.addMergedRegion(new CellRangeAddress(nrow-1,nrow,cc,cc));
									 }
									 
									 //璁剧疆琛岃竟璺濇牸寮�
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
		
		
		//淇濆瓨鏁版嵁
		
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

	    //闁藉牆顕稉宥呮倱濞村繗顫嶉崳銊︽暭閸欐绱惍锟�	     
			//閼惧嘲褰嘽onntext
			
			//鐠佸墽鐤嗛弬鍥︽mimeType
			
			String mimetype=con.getMimeType(file);
			response.setContentType(mimetype);
			//鐠佸墽鐤嗘稉瀣祰婢剁繝淇婇幁锟�			response.setHeader("content-disposition", "attchment;filename="+file);
			//鐎佃瀚瑰ù锟�			//閼惧嘲褰囨潏鎾冲弳濞达拷
			
			InputStream is=con.getResourceAsStream("/download/"+file);
			
			//閼惧嘲褰囨潏鎾冲毉濞达拷
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
