package cn.hzg.pojo;

public class plate {
@Override
	public String toString() {
		return "plate [CAS=" + CAS + ", Compound=" + Compound + ", Plate="
				+ Plate + "]";
	}
private String CAS;
private String Compound;
private String Plate;
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
	


}
