package com.hzg.services;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.hzg.domain.inExcel;
import com.hzg.domain.plate;
import com.hzg.utils.ExcelUtils;
import com.hzg.utils.getuuid;

public class ExcelServices {

	public inExcel readExcel(List<FileItem> list,String savePath) {
		// TODO Auto-generated method stub
		
		inExcel xls=new inExcel();
		File file = new File(savePath);
		int prows=0;
		int pcols=0;
		int margin_left=0;
		int margin_right=0;
		int margin_top=0;
		int margin_butto=0;
		
		List<plate> rds=null;
		//鍒ゆ柇涓婁紶鏂囦欢鐨勪繚瀛樼洰褰曟槸鍚﹀瓨鍦�
		if (!file.exists() && !file.isDirectory()) {
			
			//鍒涘缓鐩綍
			file.mkdir();
		}
		//娑堟伅鎻愮ず
		String message = "";
		try{
			//浣跨敤Apache鏂囦欢涓婁紶缁勪欢澶勭悊鏂囦欢涓婁紶姝ラ锛�
			//1銆佸垱寤轰竴涓狣iskFileItemFactory宸ュ巶
			DiskFileItemFactory factory = new DiskFileItemFactory();
			//2銆佸垱寤轰竴涓枃浠朵笂浼犺В鏋愬櫒
			ServletFileUpload upload = new ServletFileUpload(factory);
			//瑙ｅ喅涓婁紶鏂囦欢鍚嶇殑涓枃涔辩爜
			upload.setHeaderEncoding("UTF-8"); 
			//3銆佸垽鏂彁浜や笂鏉ョ殑鏁版嵁鏄惁鏄笂浼犺〃鍗曠殑鏁版嵁

			
			
			//4銆佷娇鐢⊿ervletFileUpload瑙ｆ瀽鍣ㄨВ鏋愪笂浼犳暟鎹紝瑙ｆ瀽缁撴灉杩斿洖鐨勬槸涓�釜List<FileItem>闆嗗悎锛屾瘡涓�釜FileItem瀵瑰簲涓�釜Form琛ㄥ崟鐨勮緭鍏ラ」
			
			for(FileItem item : list){
				//濡傛灉fileitem涓皝瑁呯殑鏄櫘閫氳緭鍏ラ」鐨勬暟鎹�
				if(item!=null && item.isFormField()){
					
					String name = item.getFieldName();
					
					switch (name) {
					
					case "prows":
						
						prows=Integer.valueOf(item.getString("UTF-8"));
						break;
					case "pcols":
					pcols=Integer.valueOf(item.getString("UTF-8"));
		
					break;
					case "margin_left":
					margin_left=Integer.valueOf(item.getString("UTF-8"));
					break;
					case "margin_right":
					margin_right=Integer.valueOf(item.getString("UTF-8"));
					break;
					case "margin_top":
					margin_top=Integer.valueOf(item.getString("UTF-8"));
					break;
					case "margin_butto":
					margin_butto=Integer.valueOf(item.getString("UTF-8"));
					break;
					case "":
					break;
					default:
						break;
					
					}
					
					
					//瑙ｅ喅鏅�杈撳叆椤圭殑鏁版嵁鐨勪腑鏂囦贡鐮侀棶棰�
					
					//value = new String(value.getBytes("iso8859-1"),"UTF-8");
					
				}else{//濡傛灉fileitem涓皝瑁呯殑鏄笂浼犳枃浠�

					//寰楀埌涓婁紶鐨勬枃浠跺悕绉帮紝
					String filename = item.getName();
					
					if(filename==null || filename.trim().equals("")){
						continue;
					}
					//娉ㄦ剰锛氫笉鍚岀殑娴忚鍣ㄦ彁浜ょ殑鏂囦欢鍚嶆槸涓嶄竴鏍风殑锛屾湁浜涙祻瑙堝櫒鎻愪氦涓婃潵鐨勬枃浠跺悕鏄甫鏈夎矾寰勭殑锛屽锛� c:\a\b\1.txt锛岃�鏈変簺鍙槸鍗曠函鐨勬枃浠跺悕锛屽锛�.txt
					//澶勭悊鑾峰彇鍒扮殑涓婁紶鏂囦欢鐨勬枃浠跺悕鐨勮矾寰勯儴鍒嗭紝鍙繚鐣欐枃浠跺悕閮ㄥ垎
					
					String lfilename=filename.substring(filename.lastIndexOf("."));
			
					//鑾峰彇item涓殑涓婁紶鏂囦欢鐨勮緭鍏ユ祦
					InputStream in = item.getInputStream();
					//鍒涘缓涓�釜鏂囦欢杈撳嚭娴�
					String newfile=getuuid.getUUID();
					FileOutputStream out = new FileOutputStream(savePath + "\\" +newfile+lfilename );
					//鍒涘缓涓�釜缂撳啿鍖�
					byte buffer[] = new byte[1024];
					//鍒ゆ柇杈撳叆娴佷腑鐨勬暟鎹槸鍚﹀凡缁忚瀹岀殑鏍囪瘑
					int len = 0;
					//寰幆灏嗚緭鍏ユ祦璇诲叆鍒扮紦鍐插尯褰撲腑锛�len=in.read(buffer))>0灏辫〃绀篿n閲岄潰杩樻湁鏁版嵁
					while((len=in.read(buffer))>0){
						//浣跨敤FileOutputStream杈撳嚭娴佸皢缂撳啿鍖虹殑鏁版嵁鍐欏叆鍒版寚瀹氱殑鐩綍(savePath + "\\" + filename)褰撲腑
						out.write(buffer, 0, len);
					}
					String ff=savePath+"\\"+newfile+lfilename;					
					out.flush();
					rds=new ExcelUtils().excelToList(ff);
					
					in.close();
				
					out.close();
					
					item.delete();
					
					//开启多线程，删除文件；
					 Thread thread = new FileDelete(ff);
					 thread.start();
					
					
			
					
					
					
				}
			}
		}catch (Exception e) {
			e.printStackTrace();
			return null;
		}
		
		xls.setCols(pcols);
		xls.setRows(prows);
		xls.setMargin_butto(margin_butto);
		xls.setMargin_top(margin_top);
		xls.setMargin_left(margin_left);
		xls.setMargin_right(margin_right);
		
		xls.setPlates(rds);
		return xls;
		
		
	}

	public String toExcel(HttpServletRequest request,inExcel xls) {
		
		// TODO Auto-generated method stub
		
		XSSFWorkbook book	=new ExcelUtils().getXLSXBook(request.getRealPath("/download")+"/template.xlsx");
   
		String rowxl="abcdefghijklmnopqrstuvwxyz";
		XSSFSheet sheet=book.getSheetAt(0);
		XSSFRow row=null;
		XSSFRow row1=null;
		XSSFRow trow=null;
		XSSFCell tcell=null;
		int lmar=xls.getMargin_left();
		int rmar=xls.getMargin_right();
		int tmar=xls.getMargin_top();
		int bmar=xls.getMargin_butto();
		int cols=xls.getCols();
		int rows=xls.getRows();
		int ncol=1;
		int nrow=0;
		int inrow=0;
		int btrow=11;//标题行
		int rounds=0;
		int topjjrow=1;//上行行间距行
		int bottojjrow=1;//下行间距行
		int jrr=topjjrow+2;
		XSSFFont fonta =book.createFont();
		XSSFFont fontb =book.createFont();
		XSSFFont fontc =book.createFont();
		XSSFFont fontd =book.createFont();
		XSSFFont fonte =book.createFont();
		XSSFFont fontf =book.createFont();
		XSSFCellStyle empty_cs=(sheet.getRow(11).getCell(0).getCellStyle());
		XSSFCellStyle data_cs=(sheet.getRow(11).getCell(1).getCellStyle());
		
		short EMPYT_BGK=64;
		Boolean nomarg=false;
		List<plate> list=xls.getPlates();
		
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
							System.out.println("dddd");
							sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonta,bqStyle,2,0 ,1,12));
						}
						
						
						tcell.setCellValue("Plate layout:"+list.get(listn).getPlate());
						trow.setHeight((short) (13.5*20));
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
