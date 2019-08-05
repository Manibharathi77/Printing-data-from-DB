package com.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Exceltodb {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Class.forName("com.mysql.jdbc.Driver");
		Connection con1 = DriverManager.getConnection("jdbc:mysql://localhost:3306/demo", "root", "password");
		Statement st1 = con1.createStatement();
		String qry1 = "select * from party";
		ResultSet rs = st1.executeQuery(qry1);

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("Party names");
		HSSFRow row = sheet.createRow(1);
		HSSFCell cell;
		cell = row.createCell(1);
		cell.setCellValue("party");
		cell = row.createCell(2);
		cell.setCellValue("count");
		int i = 2;
		while (rs.next()) {
			row = sheet.createRow(i);
			cell = row.createCell(1);
			cell.setCellValue(rs.getString("party"));
			cell = row.createCell(2);
			cell.setCellValue(rs.getString("count"));
			i++;
		}
		FileOutputStream out = new FileOutputStream(new File("exceldatabase(Election result).xlsx"));
		wb.write(out);
		out.close();
		System.out.println("Completed as expected");
		
	}

}
