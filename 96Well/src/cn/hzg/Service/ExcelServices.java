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
import org.junit.internal.runners.model.EachTestNotifier;
import org.springframework.web.multipart.MultipartFile;
import cn.hzg.pojo.DataInfo;
import cn.hzg.pojo.plate;
import cn.hzg.Utils.ExcelUtils;
import cn.hzg.Utils.getuuid;;

public class ExcelServices {

	public DataInfo readExcel(MultipartFile file,String savePath,DataInfo df) {
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
			
			df=new ExcelUtils().excelToList(ff,df);

			
			Thread thread = new FileDelete(ff);
			thread.start();	
			return df;
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
	
		int mv=6;

		int mvv=(mv>0?1:0);
		
		int cols=df.getCols();
		int rows=df.getRows();
		int nrow=0;
		int topjjrow=1;//涓婅琛岄棿璺濊
		int bottojjrow=1;//涓嬭闂磋窛琛�
		int jrr=topjjrow+2;
		XSSFFont fonta =book.createFont();
		XSSFFont fontb =book.createFont();
		XSSFFont fontd =book.createFont();
		XSSFFont fonte =book.createFont();
		XSSFFont fontf =book.createFont();
		XSSFCellStyle empty1_cs=(sheet.getRow(13).getCell(1).getCellStyle());
		XSSFCellStyle empty2_cs=(sheet.getRow(13).getCell(3).getCellStyle());
		XSSFCellStyle empty3_cs=(sheet.getRow(13).getCell(2).getCellStyle());
		XSSFCellStyle empty4_cs=(sheet.getRow(13).getCell(4).getCellStyle());

		XSSFCellStyle data1_cs1=(sheet.getRow(11).getCell(1).getCellStyle());
		XSSFCellStyle data1_cs2=(sheet.getRow(12).getCell(1).getCellStyle());
		XSSFCellStyle data2_cs1=(sheet.getRow(11).getCell(2).getCellStyle());
		XSSFCellStyle data2_cs2=(sheet.getRow(12).getCell(2).getCellStyle());
		XSSFCellStyle data3_cs1=(sheet.getRow(11).getCell(3).getCellStyle());
		XSSFCellStyle data3_cs2=(sheet.getRow(12).getCell(3).getCellStyle());
		
		XSSFCellStyle plateStyle=(sheet.getRow(13).getCell(0).getCellStyle());

		
		Boolean nomarg=false;
		
		List<plate> list=df.getList();
		
		int zzrow=list.size();
		int listn=0;
		int btrow=11;//从出去标题行的起始行

		int rounds=df.getRounds();
		
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
							sheet.getRow(nrow).getCell(cc).setCellStyle(plateStyle);
						}else
						{
							sheet.getRow(nrow).getCell(cc).setCellStyle(plateStyle);
						}
						//填充Plate layout
						tcell.setCellValue("Plate layout:"+list.get(listn).getPlate());
						trow.setHeight((short) (17.5*20));
						if(cc==cols){
						sheet.addMergedRegion(new CellRangeAddress(nrow,nrow,0,cols));
						}
						}
					
					//设置plate下空行
					if(topjjrow>0 && rr==topjjrow && cc==cols){
						sheet.addMergedRegion(new CellRangeAddress(nrow,nrow,0,cols));
						trow.setHeight((short) (12.75*20));
					}
					
					//设置标题顶端行格式
					if(rr==topjjrow && topjjrow>0){
						sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fonta,book.createCellStyle(),2,0 ,1,12));
					}
					//设置列标题
					if(rr==topjjrow+1){	
						//填充列标题
						trow.setHeight((short) (19.5*20));
						if(cc>0){
							 tcell.setCellValue(cc);
						}
						//设置列标题格式
						if(cc<cols){
						 //列标题除最后一个格式
						sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontb, book.createCellStyle(),0,2 ,1,10));
						}
						else
						{
						 //列标题最后一个格式
							sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontb, book.createCellStyle(),4,2 ,1,10));
						}
					}
					if(rr>topjjrow+1){
						//设置行标题
						if(cc==0){
							if((rr-jrr)%2==0){
							
							sheet.addMergedRegion(new CellRangeAddress(nrow,nrow+1,0,0));
							tcell.setCellValue(rowxl.substring((rr-jrr)/2, (rr-jrr)/2+1));
							}
							sheet.getRow(nrow).getCell(cc).setCellStyle(ExcelUtils.excelStyle(fontb, book.createCellStyle(),4,2 ,1,10));
						 
						}else{
							if((rr-jrr)%2>0 ){

										String cas=list.get(listn).getCAS().trim();
										String compound=list.get(listn).getCompound();
												
										//填充CAS
										sheet.getRow(nrow-1).getCell(cc).setCellValue(cas); 
										
										//填充Compound
										sheet.getRow(nrow).getCell(cc).setCellValue(compound);
					
										//EMPTY格式
										if(cas=="Empty"){
											 sheet.addMergedRegion(new CellRangeAddress(nrow-1,nrow,cc,cc));
											 if(cc==1 ){
												 if(listn+1<list.size()&&list.get(listn+1).getCAS().trim().equals("Empty"))
												 {
													 sheet.getRow(nrow-1).getCell(cc).setCellStyle(empty4_cs);
													 sheet.getRow(nrow).getCell(cc).setCellStyle(empty4_cs);
												 }else{
													 sheet.getRow(nrow-1).getCell(cc).setCellStyle(empty1_cs);
													 sheet.getRow(nrow).getCell(cc).setCellStyle(empty1_cs);
												 }
											 
											 }else if(cc==cols){
												 sheet.getRow(nrow-1).getCell(cc).setCellStyle(empty2_cs);
												 sheet.getRow(nrow).getCell(cc).setCellStyle(empty2_cs);

											 }else if(listn+1<list.size() && list.get(listn+1).getCAS().trim().equals("Empty")){
												 sheet.getRow(nrow-1).getCell(cc).setCellStyle(empty4_cs);
												 sheet.getRow(nrow).getCell(cc).setCellStyle(empty4_cs);
												 
											 }else{
												 sheet.getRow(nrow-1).getCell(cc).setCellStyle(empty3_cs);
												 sheet.getRow(nrow).getCell(cc).setCellStyle(empty3_cs);
											 }
											 

										}	
										//非EMPTY格式
										else{
											sheet.getRow(nrow-1).getCell(cc).setCellStyle(data1_cs1);
											 sheet.getRow(nrow).getCell(cc).setCellStyle(data1_cs2);
											 
											 if(listn>0&&list.get(listn-1).getCAS().trim().equals("Empty")){

												 sheet.getRow(nrow-1).getCell(cc).setCellStyle(data1_cs1);
												 sheet.getRow(nrow).getCell(cc).setCellStyle(data1_cs2);
											 }
											 if(listn+1<list.size() &&list.get(listn+1).getCAS().trim().equals("Empty")){

												 sheet.getRow(nrow-1).getCell(cc).setCellStyle(data1_cs1);
												 sheet.getRow(nrow).getCell(cc).setCellStyle(data1_cs2);
											 }
											 	
											 sheet.setColumnWidth(cc, (int)11.5*252+323);
										}
										
										if(listn<zzrow-1){
											listn++;
										}
										

							}
							
						}
					}	
				}
			
				
					
			}
			
		}
		
		
		
		//保存文件
		
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
