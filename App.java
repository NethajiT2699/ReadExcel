package com.xssf.readingexcel.ReadingExcel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.xssf.readingexcel.empdata.EmpData;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        try {
			XSSFWorkbook wb = new XSSFWorkbook("./data/empDetails.xlsx");
			XSSFSheet sheet=wb.getSheetAt(0);
			XSSFRow row=sheet.getRow(1);
	        int colCount=row.getLastCellNum();
	        int rowCount=sheet.getLastRowNum();
	        EmpData e1=new EmpData();
	        EmpData e2=new EmpData();
	        EmpData e3=new EmpData();
	        EmpData e4=new EmpData();
	        EmpData e5=new EmpData();
	        EmpData e6=new EmpData();
	        EmpData e7=new EmpData();
	        EmpData e8=new EmpData();
	        EmpData e9=new EmpData();
	        EmpData e10=new EmpData();
	        
	        EmpData[] emp= {e1,e2,e3,e4,e5,e6,e7,e8,e9,e10};
	        ArrayList<EmpData> al=new ArrayList<>();
	        int j=1;
	        for (int i = 0; i <emp.length; i++) {
	    
			    emp[i].setEmp_id((int)sheet.getRow(j).getCell(0).getNumericCellValue());
				emp[i].setEmp_name(sheet.getRow(j).getCell(1).getStringCellValue());
				emp[i].setAge((int)sheet.getRow(j).getCell(2).getNumericCellValue());
				emp[i].setGender(sheet.getRow(j).getCell(3).getStringCellValue());
				emp[i].setLocation(sheet.getRow(j).getCell(4).getStringCellValue());
				al.add(emp[i]);
				j++;
				System.out.println();
				}
//	        for(EmpData e:al) {
//	        	System.out.println(e);
//	        }
	        Iterator<EmpData> itr=al.iterator();
	        while(itr.hasNext()) {
	        	System.out.println(itr.next());
	        }
	        
	        
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        System.out.println("finished...!");
    }
}
