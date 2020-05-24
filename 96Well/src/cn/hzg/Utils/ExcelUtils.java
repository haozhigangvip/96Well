package cn.hzg.Utils;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.hzg.pojo.DataInfo;
import cn.hzg.pojo.plate;


public class ExcelUtils {

public DataInfo  excelToList(String filePath,DataInfo df) throws EncryptedDocumentException, InvalidFormatException
{

	Workbook book=null;
	Sheet sheet=null;
	int trows=df.getRows();
	int tcols=df.getCols();
	Row row=null;

	
	List<plate> list=new ArrayList<plate>();
	
	
	try {
		

		
		book=WorkbookFactory.create(new FileInputStream(new File(filePath)));
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
		try {
			book.close();
		} catch (IOException e1) {
			// TODO 鑷姩鐢熸垚鐨� catch 鍧�
			e1.printStackTrace();
		}
		return null;
	} 
	if(book!=null){
		sheet=book.getSheetAt(0);
		if(sheet==null){
			try {
				book.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return null;
			
		}
		int rows=sheet.getLastRowNum();//上传文件列表总行数
		int rr=1; //上传文件列表遍历当前行
		String rowname=null;//上传文件所在行
		int  colnum =0; //上传文件所在列
		String cas =null;
		String compound =null;
		String plate=null ;
		String rowarr="abcdefghijklmnopqrstuvwxyz";
		String nrowname=null;
		int ncolnum=0;
		int round=0;

		while (rr<=rows){
			round++;
			for(int arr=1;arr<=trows;arr++){
				for(int acc=1; acc<=tcols;acc++){
					row=sheet.getRow(rr);
					//计算当前遍历行标签
					nrowname=rowarr.substring(arr-1,arr);

					if(row!=null){
						//获取列标签
						colnum=(int)row.getCell(4).getNumericCellValue();
						//获取行标签
						rowname=row.getCell(3).getStringCellValue().trim();
												
						
						//判断是否为EMPTY
						if(rowname.equalsIgnoreCase(nrowname) && colnum==acc ){
							//获取CAS判断CAS数据类型，并校正为字符型
							Cell deliveryTimeCell =row.getCell(0);
			                if(deliveryTimeCell.getCellTypeEnum()== CellType.NUMERIC ){
			                	if(HSSFDateUtil.isCellDateFormatted(deliveryTimeCell)){
			                		cas=((row.getCell(0).getDateCellValue()).toLocaleString().replace("0:00:00", "").trim());
			                	}else
			                	{
			                		DecimalFormat decimalFormat = new DecimalFormat("###################.###########");
			                		cas=(decimalFormat.format(row.getCell(0).getNumericCellValue()).toString());
			                	}
			                }else{
			                		cas=(row.getCell(0).getStringCellValue());
			                	}

							//获取Compound
			                compound=row.getCell(1).getStringCellValue();
			                //获取Plate
							plate=row.getCell(2).getStringCellValue();
			
							rr++;	
						}else{
							cas="Empty";
							compound=null;
							plate=row.getCell(2).getStringCellValue().trim();
						
						}
						
					}else{
						cas="Empty";
						compound=null;
						plate=null;
						rr++;
					}
					plate pl=new plate();
					pl.setCAS(cas);
					pl.setCompound(compound);
					pl.setPlate(plate);
					pl.setRow(nrowname);
					pl.setCol(String.valueOf(acc));
					list.add(pl);

				}
			}
		}
		df.setRounds(round);
		df.setList(list);
		
		
	}
	try {
		book.close();
	
	} catch (IOException e) {
		// TODO 鑷姩鐢熸垚鐨� catch 鍧�
		e.printStackTrace();
	}
	
	return df;
}

public XSSFWorkbook getXLSXBook(String filePath ){
	XSSFWorkbook book=null;
	FileInputStream ff=null;
	try {
		ff=new FileInputStream(filePath);
		
		book = new XSSFWorkbook(ff);
	} catch (Exception e1) {
		// TODO 閼奉亜濮╅悽鐔稿灇閻拷catch 閸э拷
		e1.printStackTrace();
		return null;
		
	} finally{
		try {
			
			ff.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	return book;
	
}

@SuppressWarnings( "deprecation" )
public static CellStyle excelStyle(XSSFFont font ,CellStyle style,int Border,int Center,int fontBold,int fontSize){

	switch (Border) {
	case 0:
		
	break;
	case 1:
		style.setBorderTop(CellStyle.BORDER_THIN);
		break;
	case 2:
		style.setBorderBottom(CellStyle.BORDER_THIN);
		break;
	case 3:
		style.setBorderLeft(CellStyle.BORDER_THIN);
		break;
	case 4:
		style.setBorderRight(CellStyle.BORDER_THIN);
		break;
	case 5:
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		break;
	case 6:
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		break;	
	case 7:
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		break;
	case 8:
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_DASHED);
		break;
	case 9:
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_DASHED);	
		break;
	case 10:
		style.setBorderTop(CellStyle.BORDER_DASHED);
		style.setBorderRight(CellStyle.BORDER_THIN);
		break;

	case 11:
		style.setBorderBottom(CellStyle.BORDER_DASHED);
		style.setBorderRight(CellStyle.BORDER_THIN);

		break;
	case 12:
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		break;
	case 13:
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		break;
	case 14:

		style.setBorderTop(CellStyle.BORDER_DASHED);
		break;
	case 15:
		
		style.setBorderBottom(CellStyle.BORDER_DASHED);
		break;
	case 16:
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_DASHED);
		style.setBorderRight(CellStyle.BORDER_DASHED);
		break;
	default:
		break;
	}
	switch (Center) {
	case 0:
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
		break;
	case 1:
		style.setAlignment(CellStyle.ALIGN_CENTER);
		
		break;
	case 2:
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		break;
	default:
		break;
	}
	
	
	
		font.setFontHeightInPoints((short)(fontSize));
		font.setFontName("Arial");	
		style.setWrapText(true);
		if(fontBold!=0){
		font.setBold(true);
		}
		style.setFont(font);	

	
	
	 return style;
	
}


} 

