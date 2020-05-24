package cn.hzg.pojo;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;

public class DataInfo {
private Sheet sheet;
private int rows;
private int cols;
private int rounds;

private List<plate> list;

public Sheet getSheet() {
	return sheet;
}
public void setSheet(Sheet sheet) {
	this.sheet = sheet;
}



	
public int getRounds() {
	return rounds;
}
public void setRounds(int rounds) {
	this.rounds = rounds;
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
