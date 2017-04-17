package com.XLexample.Actions;

import java.io.File;

import org.apache.struts2.ServletActionContext;

import com.XLexample.XLSOperations.XLSReadWriter;

public class XLSActions {
	private String strMessage;
	private String uploadFileStrName;
	
	
	public String uploadXLData()
	{
		String filePath = ServletActionContext.getServletContext().getRealPath("/");
		System.out.println(filePath + uploadFileStrName);
		File inputFile = new File(filePath + uploadFileStrName);
		File outputFile = new File("E:\\Test1.csv");
		XLSReadWriter.uploadXLS(inputFile, outputFile);
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
