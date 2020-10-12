package step_definition;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;

public class StepDefinition {
	
	String excel_path="C:\\Users\\VENKATRAMAN\\workspace\\ExcelManipulation\\ExcelManipulation\\src\\main\\java"
			+ "\\test_data\\sample_excelfile.xlsx";
	String new_excel_path="C:\\Users\\VENKATRAMAN\\workspace\\ExcelManipulation\\ExcelManipulation\\src\\main\\java"
			+ "\\test_data\\output_excelfile.xlsx";
	
	public Set<String> duplicate_finder;
	public File excel;
	public File new_excel;
	public FileInputStream fis_excel;
	public FileInputStream fis_new_excel;
	public XSSFWorkbook excel_wb;
	public XSSFWorkbook new_excel_wb;
	public XSSFSheet excel_sheet;
	public XSSFSheet new_excel_sheet;
	public CellStyle excel_style;
	public CellStyle new_excel_style;
	public Row excel_sheet_row;
	public Cell excel_sheet_cell;
	public Row new_excel_sheet_row;
	public Cell new_excel_sheet_cell;
	public FileOutputStream fout_new_excel;
	public String cell_value="temp";
	
	


@Given("^Excel file value copy and fetch in new Excel File$")
public void excel_file_value_copy_and_fetch_in_new_Excel_File() throws Exception  {
   
	//path assign
	excel=new File(excel_path);
	new_excel=new File(new_excel_path);
	fis_excel=new FileInputStream(excel);
	fis_new_excel=new FileInputStream(new_excel);
	excel_wb=new XSSFWorkbook(fis_excel);
	new_excel_wb=new XSSFWorkbook(fis_new_excel);
	excel_sheet=excel_wb.getSheetAt(0);
	new_excel_sheet=new_excel_wb.createSheet();
	excel_style=excel_wb.getCellStyleAt(2);
	new_excel_style=new_excel_wb.createCellStyle();
	
	
	//header
	for(int i=0;i<excel_sheet.getRow(0).getLastCellNum();i++) {
		
		excel_sheet_row=excel_sheet.getRow(0);
		excel_sheet_cell=excel_sheet_row.getCell(i);
		cell_value=excel_sheet_cell.getStringCellValue();
		
		if(i==0)new_excel_sheet_row=new_excel_sheet.createRow(0);
		else new_excel_sheet_row=new_excel_sheet.getRow(0);
		new_excel_sheet_cell=new_excel_sheet_row.createCell(i);
		new_excel_sheet_cell.setCellValue(cell_value);
		if(i==9) new_excel_sheet_cell.setCellValue("Time Ingested (AEST Time)");
		new_excel_style.cloneStyleFrom(excel_style);
		new_excel_sheet_cell.setCellStyle(new_excel_style);
	}
	
	//value
	
		
		
	for(int i=1;i<excel_sheet.getLastRowNum();i++) {
			new_excel_sheet_row=new_excel_sheet.createRow(i);
		for(int j=0;j<excel_sheet.getRow(i).getLastCellNum();j++) {
			
			
			//get numberic cell value convert into date(String)
			if(j==1||j==7) {
			excel_sheet_row=excel_sheet.getRow(i);
			excel_sheet_cell=excel_sheet_row.getCell(j);
			double date=excel_sheet_cell.getNumericCellValue(); 
			Date dateformat=DateUtil.getJavaDate(date);
			String value=new SimpleDateFormat("dd-MM-yyyy").format(dateformat);
			new_excel_sheet_row=new_excel_sheet.getRow(i);
			new_excel_sheet_cell=new_excel_sheet_row.createCell(j);
			new_excel_sheet_cell.setCellValue(value);
			}
			
			//get numeric cell value convert into time(string)
			else if(j==2||j==9) {
			excel_sheet_row=excel_sheet.getRow(i);
			excel_sheet_cell=excel_sheet_row.getCell(j);
			double time=excel_sheet_cell.getNumericCellValue();
			Date timeformat=DateUtil.getJavaDate(time);
			String value=new SimpleDateFormat("h:mm a").format(timeformat);
			new_excel_sheet_row=new_excel_sheet.getRow(i);
			new_excel_sheet_cell=new_excel_sheet_row.createCell(j);
			new_excel_sheet_cell.setCellValue(value);
			}
			
			//copy and fetching string value 
			
			else {
				excel_sheet_row=excel_sheet.getRow(i);
				excel_sheet_cell=excel_sheet_row.getCell(j);
				String value=excel_sheet_cell.getStringCellValue();
				new_excel_sheet_row=new_excel_sheet.getRow(i);
				new_excel_sheet_cell=new_excel_sheet_row.createCell(j);
				new_excel_sheet_cell.setCellValue(value);	
			}
		}
	}  
}

@Then("^Convert MNL to AEST$")
public void convert_MNL_to_AEST() throws Exception {
	
	//convert MNL to AEST
	
    for(int i=1;i<excel_sheet.getLastRowNum();i++) {
    	excel_sheet_row=excel_sheet.getRow(i);
		excel_sheet_cell=excel_sheet_row.getCell(9);
		double time=excel_sheet_cell.getNumericCellValue();
		Date timeformat=DateUtil.getJavaDate(time);
		String value=new SimpleDateFormat("hh:mm:ss a").format(timeformat);
		
		SimpleDateFormat format1=new SimpleDateFormat("hh:mm:ss a");
		
		Date time_1=format1.parse(value);
		long update_time=time_1.getTime()+2*1000*60*60;      //Australian Eastern Standard Time is 2 hours ahead of Manila, Metro Manila, Philippines 
		String outputvalue=new SimpleDateFormat("h:mm a").format(update_time);
		
		new_excel_sheet_row=new_excel_sheet.getRow(i);
		new_excel_sheet_cell=new_excel_sheet_row.createCell(9);
		new_excel_sheet_cell.setCellValue(outputvalue);
	
    }
    
}

@And("^Create new Column and find time taken$")
public void create_new_Column_and_find_time_taken() throws Exception {
	
	
	//time taken column
	
	
	new_excel_sheet_row=new_excel_sheet.getRow(0);
	new_excel_sheet_cell=new_excel_sheet_row.createCell(12);
	new_excel_sheet_cell.setCellValue("Time Taken(Days)");
	new_excel_style.cloneStyleFrom(excel_style);
	new_excel_sheet_cell.setCellStyle(new_excel_style);
	
	for(int i=1;i<excel_sheet.getLastRowNum();i++) {
		excel_sheet_row=excel_sheet.getRow(i);
		Cell excel_sheet_cell_1=excel_sheet_row.getCell(1);
		Cell excel_sheet_cell_7=excel_sheet_row.getCell(7);
		
		double date_Received=excel_sheet_cell_1.getNumericCellValue(); 
		double date_Decision=excel_sheet_cell_7.getNumericCellValue();
		Date dateformat_cell_1=DateUtil.getJavaDate(date_Received);
		Date dateformat_cell_7=DateUtil.getJavaDate(date_Decision);
		String value_cell_1=new SimpleDateFormat("dd-MM-yyyy").format(dateformat_cell_1);
		String value_cell_7=new SimpleDateFormat("dd-MM-yyyy").format(dateformat_cell_7);
		SimpleDateFormat format=new SimpleDateFormat("dd-MM-yyyy");
		Date d1=format.parse(value_cell_1);
		Date d2=format.parse(value_cell_7);
		long difference=(d1.getTime()-d2.getTime())/(1000*60*60*24);
		new_excel_sheet_row=new_excel_sheet.getRow(i);
		new_excel_sheet_cell=new_excel_sheet_row.createCell(12);
		new_excel_sheet_cell.setCellValue(difference);
	}
	
    
   
}

@Then("^Remove Dupliacte VCI-Codes and their respective Rows$")
public void remove_Dupliacte_VCI_Codes_and_their_respective_Rows() throws Exception{
	
	
			duplicate_finder=new HashSet<String>();
		//duplicate Remove
		for(int i=1;i<new_excel_sheet.getLastRowNum();i++) {
			if(!duplicate_finder.add(new_excel_sheet.getRow(i).getCell(0).getStringCellValue())) {
				new_excel_sheet.removeRow(new_excel_sheet.getRow(i));     //remove row
			}
		}
		
		
		//removing blank rows
		for(int i = 1; i < new_excel_sheet.getLastRowNum(); i++){
			
		    if(new_excel_sheet.getRow(i)==null){
	    	new_excel_sheet.shiftRows(i + 1, new_excel_sheet.getLastRowNum(), -1);
		        i--;
		    }
		}
		for(int i=0;i<new_excel_sheet.getRow(0).getLastCellNum();i++) {
		new_excel_sheet.autoSizeColumn(i);
}
	
	
	
    fout_new_excel=new FileOutputStream(new_excel);
    new_excel_wb.write(fout_new_excel);
   
}


}
