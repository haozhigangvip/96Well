package com.hzg.domain;

import java.util.List;

public class inExcel {

private int rows;
private int cols;
private int margin_left;
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
private int margin_right;
private int margin_top;
private int margin_butto;
private List<plate> plates;
	
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
public List<plate> getPlates() {
	return plates;
}
public void setPlates(List<plate> plates) {
	this.plates = plates;
}
public boolean isReadError(){
	boolean err= false;
	if(this.getPlates()==null || this.getCols()==0 || this.getRows()==0){
		err=true;
	}
	return err;
}
}
