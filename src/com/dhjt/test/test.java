package com.dhjt.test;

import com.dhjt.office2html.POIWordToHtml;

/**
 * 测试类
 * @author DHJT 2018年5月1日 下午1:04:20
 *
 */
public class test {

	public static void main(String[] args) {
		POIWordToHtml.wordToHtml("C:/Workspaces/TestTemp/test.docx", "C:/Workspaces/TestTemp/poi/image", "C:/Workspaces/TestTemp/poi/html/test.html");
//		POIWordToHtml.wordToHtml("C://poi/test2.doc", "C://poi/image", "C://poi/html/test2.html");
//		POIExcelToHtml.excelToHtml("C:/Workspaces/TestTemp/2017-11-20.xlsx", "C:/Workspaces/TestTemp/poi/html/test3.html", true);
//		POIExcelToHtml.excelToHtml("C://poi/test4.xls", "C://poi/html/test4.html", true);
//		POIPptToHtml.pptToHtml("C:/Workspaces/TestTemp/商务项目企划书PPT模板2.pptx", "C:/Workspaces/TestTemp/poi/html/ppt");
//		POIPptToHtml.pptToHtml("C://poi/test6.ppt", "C://poi/html/ppt");
	}
}
