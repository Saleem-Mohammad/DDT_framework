package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtil {
	XSSFWorkbook wb;
	//constructor for reading excel path
	public ExcelFileUtil(String excelpath) throws Throwable 
	{
		FileInputStream fi=new FileInputStream(excelpath);
		wb=new XSSFWorkbook(fi);
	}
	//method for row count
	public int rowCount(String sheetname) 
	{
		return wb.getSheet(sheetname).getLastRowNum();
	}
	//method to count no of cell in a row
	public int cellCount(String sheetname) 
	{
		return wb.getSheet(sheetname).getRow(0).getLastCellNum();
	}
	//method for cell data
	public String getcelldata(String sheetname,int row,int coloumn)  
	{
		String data="";
		if(wb.getSheet(sheetname).getRow(row).getCell(coloumn).getCellType()==Cell.CELL_TYPE_NUMERIC) 
		{
			int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(coloumn).getNumericCellValue();
			data=String.valueOf(celldata);
		}
		else 
		{
			data=wb.getSheet(sheetname).getRow(row).getCell(coloumn).getStringCellValue();
		}
		return data;
		
	}
	//method for set cell data
	public void setcelldata(String sheetname,int row,int coloumn,String status,String writeexcel) throws Throwable 
	{
		//get sheet from workbook
		XSSFSheet ws=wb.getSheet(sheetname);
		//get row from sheet
		XSSFRow rownum=ws.getRow(row);
		//create cell in a row
		XSSFCell cellnum=rownum.createCell(coloumn);
		//write status
		cellnum.setCellValue(status);
		if(status.equalsIgnoreCase("pass")) 
		{
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			font.setColor(IndexedColors.BRIGHT_GREEN.getIndex());
			font.setBold(true);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rownum.getCell(coloumn).setCellStyle(style);
			
		}
		else if(status.equalsIgnoreCase("fail")) 
		{
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rownum.getCell(coloumn).setCellStyle(style);
			
		}
		else if(status.equalsIgnoreCase("blocked")) 
		{
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rownum.getCell(coloumn).setCellStyle(style);
			
		}
		FileOutputStream fo=new FileOutputStream(writeexcel);
		wb.write(fo);		
	}
	public static void main(String[] args) throws Throwable {
		ExcelFileUtil sm=new ExcelFileUtil("C://New Folder//dummy.xlsx");
		//count no of rows
		int rc=sm.rowCount("Sheet1");
		//count no of cell in row
		int cc=sm.cellCount("Sheet1");
		System.out.println(rc+"  "+cc);
		for(int i=1;i<=rc;i++) 
		{
			String user=sm.getcelldata("Sheet1",i,0);
			String pass=sm.getcelldata("Sheet1",i,1);
			System.out.println(user+"  "+pass);
			//sm.setcelldata("Sheet1",i,2,"pass","C://New Folder//results.xlsx");
			//sm.setcelldata("Sheet1",i,2,"fail","C://New Folder//results.xlsx");
			sm.setcelldata("Sheet1",i,2,"blocked","C://New Folder//results.xlsx");
		}
	}

}
