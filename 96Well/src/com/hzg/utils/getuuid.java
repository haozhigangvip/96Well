package com.hzg.utils;

import java.util.UUID;

public class getuuid {
public static  String getUUID(){
	
		return UUID.randomUUID().toString().replace("-", "");
}
}
