package cn.hzg.pojo;

public class plate {

private String CAS;
private String Compound;
private String Plate;
private String Row;
private String Col;


public String getRow() {
	return Row;
}
public void setRow(String row) {
	Row = row;
}
public String getCol() {
	return Col;
}
public void setCol(String col) {
	Col = col;
}
public String getCAS() {
	return CAS;
}
public void setCAS(String cAS) {
	CAS = cAS;
}
public String getCompound() {
	return Compound;
}
public void setCompound(String compound) {
	Compound = compound;
}
public String getPlate() {
	return Plate;
}
public void setPlate(String plate) {
	Plate = plate;
}
@Override
public String toString() {
	return "plate [CAS=" + CAS + ", Compound=" + Compound + ", Plate=" + Plate + ", Row=" + Row + ", Col=" + Col + "]";
}
	


}
