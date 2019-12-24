package exc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hslf.model.Sheet;
import org.apache.poi.hssf.usermodel.examples.CellTypes;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * ���ܣ���ȡһ��xlsx�ļ����޸ĸ��ı��ַ�����д��
 * 
 * @author zz
 *
 */
public class Demo2 {

	public static void main(String[] args) throws IOException {
		//��ȡ�ļ���
		File file = new File("D:\\1.xlsx");
		FileInputStream inputStream =  new FileInputStream(file);
		//��ȡworkbook
		 XSSFWorkbook xssfWorkbook = null;
	        try {
	            xssfWorkbook = new XSSFWorkbook(inputStream);
	        } catch (Exception e) {
	            System.out.println("Excel data file cannot be found!");
	        }
	   //��ȡsheet    
	        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);        	        		    
		//To apply a single set of text formatting (colour, style, font etc) to a cell, 
        //you should create a CellStyle for the workbook, then apply to the cells.
	       String hundred_zero = "00"; //Ϊ�˴��  002 ��Ч������ƴ��
	       String ten_zero = "0";
	       String added = "0";         
	       	for( int j = 0 ; j < 400 ; j++) {  //����400*2����Ҫ�ĵ�Ԫ��
	       		for(int k = 0 ; k < 2 ; k++) {
	       			if( j < 10) {
	       				added = hundred_zero+String.valueOf(j);
	       			}
	       			else if( j < 100 && j > 9) {
	       				 added = ten_zero+String.valueOf(j);
	       			}
	       			else {
	       				 added = String.valueOf(j);
	       			}
	        //��ȡ����j,k ��ȡĳ��ĳ��ָ����Ԫ��
	        Row row = sheet.getRow(j);   
		    Cell cell = row.getCell(k);   
		    XSSFRichTextString rt = new XSSFRichTextString("\r�������� �����ͬ��\r2020��ӭ�´�����������\r");
		    //font1��ʽ�����÷�Χ��0-26���ַ���Ҳ����XSSFRichTextString�е��ַ���
		    XSSFFont font1 = xssfWorkbook.createFont();
		    //�Ƿ�Ӵ�
		    font1.setBold(true);
		    //����ʹ�С
		    font1.setFontHeightInPoints((short)16);
		    font1.setFontName("����");
		    rt.applyFont(0, 26, font1);
		    
		    XSSFFont font2 = xssfWorkbook.createFont();
		    font2.setBold(false);
		    font2.setFontHeightInPoints((short)72);		  
		    font2.setFontName("����");
			rt.append(added,font2);

		    XSSFFont font3 = xssfWorkbook.createFont();
		    font3.setBold(true);
		    font3.setFontHeightInPoints((short)12);
		    rt.append("\rHappyCoding��HelloWorld����������\r", font3);
            //��ʽ�������set��cell
            cell.setCellValue(rt);
	       		}
	       	}     
	       	sheet.setColumnWidth(0, (int) 11270);
	       	sheet.setColumnWidth(1, (int) 11270);
			try (OutputStream fileOut = new FileOutputStream("D:\\1.xlsx")) {			    
		    	xssfWorkbook.write(fileOut);	
				System.out.println("д�����");
					}
	       	
	}
}