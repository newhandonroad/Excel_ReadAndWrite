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
 * 功能：读取一个xlsx文件，修改富文本字符串后写入
 * 
 * @author zz
 *
 */
public class Demo2 {

	public static void main(String[] args) throws IOException {
		//获取文件流
		File file = new File("D:\\1.xlsx");
		FileInputStream inputStream =  new FileInputStream(file);
		//获取workbook
		 XSSFWorkbook xssfWorkbook = null;
	        try {
	            xssfWorkbook = new XSSFWorkbook(inputStream);
	        } catch (Exception e) {
	            System.out.println("Excel data file cannot be found!");
	        }
	   //获取sheet    
	        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);        	        		    
		//To apply a single set of text formatting (colour, style, font etc) to a cell, 
        //you should create a CellStyle for the workbook, then apply to the cells.
	       String hundred_zero = "00"; //为了达成  002 的效果进行拼接
	       String ten_zero = "0";
	       String added = "0";         
	       	for( int j = 0 ; j < 400 ; j++) {  //生成400*2个需要的单元格
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
	        //获取根据j,k 获取某行某列指定单元格
	        Row row = sheet.getRow(j);   
		    Cell cell = row.getCell(k);   
		    XSSFRichTextString rt = new XSSFRichTextString("\r共建共享 “临里”同乐\r2020年迎新春居民联欢会\r");
		    //font1格式，适用范围是0-26个字符，也就是XSSFRichTextString中的字符串
		    XSSFFont font1 = xssfWorkbook.createFont();
		    //是否加粗
		    font1.setBold(true);
		    //字体和大小
		    font1.setFontHeightInPoints((short)16);
		    font1.setFontName("宋体");
		    rt.applyFont(0, 26, font1);
		    
		    XSSFFont font2 = xssfWorkbook.createFont();
		    font2.setBold(false);
		    font2.setFontHeightInPoints((short)72);		  
		    font2.setFontName("宋体");
			rt.append(added,font2);

		    XSSFFont font3 = xssfWorkbook.createFont();
		    font3.setBold(true);
		    font3.setFontHeightInPoints((short)12);
		    rt.append("\rHappyCoding“HelloWorld”党建联盟\r", font3);
            //格式设置完后，set入cell
            cell.setCellValue(rt);
	       		}
	       	}     
	       	sheet.setColumnWidth(0, (int) 11270);
	       	sheet.setColumnWidth(1, (int) 11270);
			try (OutputStream fileOut = new FileOutputStream("D:\\1.xlsx")) {			    
		    	xssfWorkbook.write(fileOut);	
				System.out.println("写入完成");
					}
	       	
	}
}