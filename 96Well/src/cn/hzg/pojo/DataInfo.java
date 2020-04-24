package cn.hzg.pojo;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;

public class DataInfo {
private Sheet sheet;
private int rows;
private int cols;
private int margin_left;
private int margin_right;
private int margin_top;
private int margin_butto;
private List<plate> list;

public Sheet getSheet() {
	return sheet;
}
public void setSheet(Sheet sheet) {
	this.sheet = sheet;
}

public int getMargin_left() {
	return margin_left;
}
public void setMargin_left(int margin_left) {
	this.margin_left = margin_left;
}
public int getMargin_right() {
	return margin_right;
}
public void setMargin_right(int margin_right) {
	this.margin_right = margin_right;
}
public int getMargin_top() {
	return margin_top;
}
public void setMargin_top(int margin_top) {
	this.margin_top = margin_top;
}
public int getMargin_butto() {
	return margin_butto;
}
public void setMargin_butto(int margin_butto) {
	this.margin_butto = margin_butto;
}

	
public int getRows() {
	return rows;
}
public void setRows(int rows) {
	this.rows = rows;
}
public int getCols() {
	return cols;
}
public void setCols(int cols) {
	this.cols = cols;
}
public List<plate> getList() {
	return list;
}
public void setList(List<plate> list) {
	this.list = list;
}
public boolean isReadError(){
	boolean err= false;
	if(this.getList()==null || this.getCols()==0 || this.getRows()==0){
		err=true;
	}
	return err;
}
}
