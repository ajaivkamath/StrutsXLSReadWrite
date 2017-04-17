package com.XLexample.Actions;

import java.io.File;
import java.util.List;

import org.apache.struts2.ServletActionContext;

import com.XLexample.XLSOperations.XLSReadWriter;

public class XLSActions {
	private String strMessage;
	private String uploadFileStrName;
	
	
	public String uploadXLData()
	{
		String filePath = ServletActionContext.getServletContext().getRealPath("/");
		System.out.println(filePath + uploadFileStrName);
//		File inputFile = new File(filePath + uploadFileStrName);
		File inputFile = new File("C:\\Users\\Ajai V Kamath\\Documents\\Blacklisted hospital.xlsx");
		File outputFile = new File("E:\\Test1.csv");
		File outputExtractFile = new File(filePath + "New_" + uploadFileStrName);
		List<Object[]> dataObjectList = XLSReadWriter.uploadXLS(inputFile, outputFile);
		XLSReadWriter.writeXLS(dataObjectList, outputExtractFile);
		
		return "success"; 
	}


	/**
	 * @return the uploadFileStrName
	 */
	public String getUploadFileStrName() {
		return uploadFileStrName;
	}


	/**
	 * @param uploadFileStrName the uploadFileStrName to set
	 */
	public void setUploadFileStrName(String uploadFileStrName) {
		this.uploadFileStrName = uploadFileStrName;
	}


	/**
	 * @return the strMessage
	 */
	public String getStrMessage() {
		return strMessage;
	}


	/**
	 * @param strMessage the strMessage to set
	 */
	public void setStrMessage(String strMessage) {
		this.strMessage = strMessage;
	}
	

}
