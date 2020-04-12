package cn.hzg.Service;

import java.io.File;
//多线程，删除临时Excel文件
public class FileDelete extends Thread {
	private String filepath;
	public FileDelete(String filepath){
		this.filepath=filepath;
		System.out.println(filepath);
	}
	public void run(){
		File ff= new File(filepath);
		while(ff.exists()){
			ff.delete();
		}
		
	}
}
