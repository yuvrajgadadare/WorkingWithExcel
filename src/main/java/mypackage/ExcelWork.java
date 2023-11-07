package mypackage;
import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
public class ExcelWork {

	public void ReadData() {
		try {
			File f=new File("E:\\files\\Students.xls");
			FileInputStream fin=new FileInputStream(f);
			HSSFWorkbook book=new HSSFWorkbook(fin);
			HSSFSheet sheet=book.getSheet("Sheet1");
			int r=sheet.getLastRowNum();
			System.out.println("Last Row Number="+r);
//			String name=sheet.getRow(2).getCell(1).getStringCellValue();
//			System.out.println(name);
			int i;
			for(i=1;i<=r;i++) {
				int rno=(int)sheet.getRow(i).getCell(0).getNumericCellValue();
				String name=sheet.getRow(i).getCell(1).getStringCellValue();
				float eng=(float)sheet.getRow(i).getCell(2).getNumericCellValue();
				float math=(float)sheet.getRow(i).getCell(3).getNumericCellValue();
				float sci=(float)sheet.getRow(i).getCell(4).getNumericCellValue();
				
				System.out.println(rno+" "+name+" "+eng+" "+math+" "+sci);
			}
			
		}
		catch(Exception ex) {
			System.out.println("Exception=>"+ex.getMessage());
		}
	}
	
	
	public void WriteData() {
		try {
			File f=new File("E:\\files\\Students.xls");
			FileInputStream fin=new FileInputStream(f);
			HSSFWorkbook book=new HSSFWorkbook(fin);
			FileOutputStream fout=new FileOutputStream(f);
			HSSFSheet sheet=book.getSheet("Sheet1");
			int r=sheet.getLastRowNum();
			//System.out.println("Last Row Number="+r);
//			String name=sheet.getRow(2).getCell(1).getStringCellValue();
//			System.out.println(name);
			int i;
			for(i=1;i<=r;i++) {
				int rno=(int)sheet.getRow(i).getCell(0).getNumericCellValue();
				String name=sheet.getRow(i).getCell(1).getStringCellValue();
				float eng=(float)sheet.getRow(i).getCell(2).getNumericCellValue();
				float math=(float)sheet.getRow(i).getCell(3).getNumericCellValue();
				float sci=(float)sheet.getRow(i).getCell(4).getNumericCellValue();
				float total=eng+math+sci;
				float per=total/3;
				String res="Fail";
				if(per>=50) {
					res="Pass";
				}
				HSSFCell celltotal=sheet.getRow(i).createCell(5);
				celltotal.setCellValue(total);
				HSSFCell cellper=sheet.getRow(i).createCell(6);
				cellper.setCellValue(per);
				HSSFCell cellres=sheet.getRow(i).createCell(7);
				cellres.setCellValue(res);
				
				
				System.out.println(rno+" "+name+" "+eng+" "+math+" "+sci+" "+total+" "+per+" "+res);
			}
			book.write(fout);
			fout.close();
			book.close();
			System.out.println("Finished");
			
			
		}
		catch(Exception ex) {
			System.out.println("Exception=>"+ex.getMessage());
		}
	}
	public static void main(String[] args) {


		ExcelWork e=new ExcelWork();
		//e.ReadData();
		e.WriteData();

	}

}
