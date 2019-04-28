package com.wiwj.cbs;  
  
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;  
import java.io.FileOutputStream;
import java.io.IOException;  
import java.io.InputStream;
import java.util.ArrayList;  
import java.util.List;  

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

import com.alibaba.fastjson.JSON;

/** 
 * excel读写工具类 
 */  
public class ExcelRead {  
    private final static String xls = "xls";  
    private final static String xlsx = "xlsx";  

    public static void readExcelTest(
    		String inFileName)
    	throws IOException {  
    	XSSFWorkbook workBook = null;
    	try {
    		InputStream inStream = new FileInputStream(inFileName);
            // 构造 XSSFWorkbook 对象，strPath 传入文件路径 
    		workBook = new XSSFWorkbook(inStream); 
            
    		XSSFSheet sheet = workBook.getSheetAt(0);
    		for(int i = 1; i <= sheet.getLastRowNum(); i++) {
	    		XSSFRow row = sheet.getRow(i);
	    		if(row == null) { continue; }
	    		// 第二列：机器码
	    		XSSFCell cellEncMachineCode = row.getCell(1);
	    		if(cellEncMachineCode == null) { continue; }
	    		if(cellEncMachineCode.getCellType() == Cell.CELL_TYPE_NUMERIC) {
	    		}
	    		// 第八列：硬盘序列号
	    		XSSFCell cellHdSn = row.getCell(7);
	    		if(cellHdSn == null) {
	    			cellHdSn = row.createCell(7);
	    		}
	    		// 第九列：硬盘序列号+硬盘型号前3位
	    		XSSFCell cellHdSnHdModelPrefix = row.getCell(8);
	    		if(cellHdSnHdModelPrefix == null) {
	    			cellHdSnHdModelPrefix = row.createCell(8);
	    		}
	    		// 第十列：路由MAC地址
	    		XSSFCell cellRouteMac = row.getCell(9);
	    		if(cellRouteMac == null) {
	    			cellRouteMac = row.createCell(9);
	    		}
	    		XSSFCell cellPlainMachineCode = row.getCell(10);
	    		if(cellPlainMachineCode == null) {
	    			cellPlainMachineCode = row.createCell(10);
	    		}
	    		
	    		String strMachineCodeEncrypted;
    			strMachineCodeEncrypted = cellEncMachineCode.getStringCellValue();
	    		if(strMachineCodeEncrypted == null) { continue; }
	    		String strMachineCode = "";
	    		
	    		String strHdSn = "";
	    		String strHdSnHdModelPrefix = "";
	    		String strRouteMac = "";
	    		
	    		if(strMachineCode.length() > 15) {
	    			strRouteMac = strMachineCode.substring(strMachineCode.length() - 12, strMachineCode.length());
	    			strHdSn = strMachineCode.substring(1, strMachineCode.length() - 15);
	    			strHdSnHdModelPrefix = strMachineCode.substring(1, strMachineCode.length() - 12);
	    		} else if(strMachineCode.length() > 3) {
	    			strRouteMac = "";
	    			strHdSn = strMachineCode.substring(1, strMachineCode.length() - 3);
	    			strHdSnHdModelPrefix = strMachineCode.substring(1, strMachineCode.length());
	    		} else {
	    			continue;
	    		}
	    			
	    		cellHdSn.setCellValue(strHdSn);
	    		cellHdSnHdModelPrefix.setCellValue(strHdSnHdModelPrefix);
	    		cellRouteMac.setCellValue(strRouteMac);
	    		cellPlainMachineCode.setCellValue(strMachineCode);
    		}    		
    	} catch (Exception e) {
    		JOptionPane.showMessageDialog(
    				null, 
    				e.getClass().toString(), 
    				"ERROR_MESSAGE",
    				JOptionPane.ERROR_MESSAGE);
    		
    		e.printStackTrace();
    	} finally {
    		if (workBook != null) {
    			workBook.close();
    		}
    	}
    }
    
    /** 
     * 读入excel文件，解析后返回 
     */  
    public static List<String[]> readExcel(String fileName) 
    		throws IOException {  
        // 获得文件
    	File file = new File(fileName);
        // 检查文件  
        checkFile(file);  
        // 获得Workbook工作薄对象  
        InputStream inStream = new FileInputStream(fileName);
        Workbook workbook = getWorkBook(inStream);  
        // 创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回  
        List<String[]> listResult = new ArrayList<String[]>();  
        if(workbook != null) {  
            for(int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {  
                // 获得当前sheet工作表  
                Sheet sheet = workbook.getSheetAt(sheetNum);  
                if(sheet == null) {  
                    continue;  
                }  
                // 获得当前sheet的开始行 
                int firstRowNum  = sheet.getFirstRowNum();  
                // 获得当前sheet的结束行 
                int lastRowNum = sheet.getLastRowNum();  
                // 循环除了第一行的所有行 
                for(int rowNum = firstRowNum+1; rowNum <= lastRowNum; rowNum++) {  
                    // 获得当前行
                    Row row = sheet.getRow(rowNum);  
                    if(row == null) {  
                        continue;  
                    }  
                    // 获得当前行的开始列  
                    int firstCellNum = row.getFirstCellNum();  
                    // 获得当前行的列数
                    int lastCellNum = row.getPhysicalNumberOfCells();  
                    String[] cells = new String[row.getPhysicalNumberOfCells()];  
                    // 循环当前行
                    for(int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {  
                        Cell cell = row.getCell(cellNum);  
                        cells[cellNum] = getCellValue(cell);  
                    }  
                    listResult.add(cells);  
                }  
            }  
            workbook.close();  
        }  
        return listResult;  
    }  
    
    public static void checkFile(File file) 
    	throws IOException{  
        // 判断文件是否存在
        if(null == file) {  
            throw new FileNotFoundException("文件不存在！");  
        }  
        // 获得文件名
        String fileName = file.getName();  
        // 判断文件是否是excel文件  
        if(!fileName.endsWith(xls) && 
        	!fileName.endsWith(xlsx)) {  
            throw new IOException(fileName + "不是excel文件");  
        }  
    }  
    
    public static Workbook getWorkBook(InputStream inputStream) {  
        // 创建Workbook工作薄对象
        Workbook workbook = null;  
        try {  
            workbook = new XSSFWorkbook(inputStream);  
        } catch (Exception e) {  
        	JOptionPane.showMessageDialog(
    				null, 
    				e.getMessage(), 
    				"错误提示",
    				JOptionPane.ERROR_MESSAGE);
        	e.printStackTrace();
        }  
        
        return workbook;  
    }  
    
    public static String getCellValue(Cell cell) {  
        String cellValue = "";  
        if(cell == null) {  
            return cellValue;  
        }  
        // 把数字当成String来读，避免出现1读成1.0的情况  
        if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {  
            cell.setCellType(Cell.CELL_TYPE_STRING);  
        }  
        // 判断数据的类型  
        switch (cell.getCellType()) {  
            case Cell.CELL_TYPE_NUMERIC: 
            	// 数字  
                cellValue = String.valueOf(cell.getNumericCellValue());  
                break;  
            case Cell.CELL_TYPE_STRING: 
            	// 字符串  
                cellValue = String.valueOf(cell.getStringCellValue());  
                break;  
            case Cell.CELL_TYPE_BOOLEAN: 
            	// Boolean  
                cellValue = String.valueOf(cell.getBooleanCellValue());  
                break;  
            case Cell.CELL_TYPE_FORMULA: 
            	// 公式  
                cellValue = String.valueOf(cell.getCellFormula());  
                break;  
            case Cell.CELL_TYPE_BLANK: 
            	// 空值   
                cellValue = "";
                break;  
            case Cell.CELL_TYPE_ERROR: 
            	// 故障  
                cellValue = "非法字符";  
                break;  
            default:  
                cellValue = "未知类型";  
                break;  
        }  
        return cellValue;  
    }
    
    /**
	 * 格式化Json串
	 * 
	 * @param jsonStr
	 * @return
	 */
	public static String formatJson(String jsonStr) {
		if (null == jsonStr || "".equals(jsonStr)) {
			return "";
		}
		StringBuilder strBuilder = new StringBuilder();
		char last = '\0';
		char current = '\0';
		int indent = 0;
		boolean isInQuotationMarks = false;
		for (int i = 0; i < jsonStr.length(); i++) {
			last = current;
			current = jsonStr.charAt(i);
			switch (current) {
			case '"':
				if (last != '\\') {
					isInQuotationMarks = !isInQuotationMarks;
				}
				strBuilder.append(current);
				break;
			case '{':
			case '[':
				strBuilder.append(current);
				if (!isInQuotationMarks) {
					strBuilder.append('\n');
					indent++;
					addIndentBlank(strBuilder, indent);
				}
				break;
			case '}':
			case ']':
				if (!isInQuotationMarks) {
					strBuilder.append('\n');
					indent--;
					addIndentBlank(strBuilder, indent);
				}
				strBuilder.append(current);
				break;
			case ',':
				strBuilder.append(current);
				if (last != '\\' && !isInQuotationMarks) {
					strBuilder.append('\n');
					addIndentBlank(strBuilder, indent);
				}
				break;
			default:
				strBuilder.append(current);
			}
		}

		return strBuilder.toString();
	}

    /**
     * 添加space
     * @param sb
     * @param indent
     */
    private static void addIndentBlank(
    		StringBuilder sb, 
    		int indent) {
        for (int i = 0; i < indent; i++) {
            sb.append('\t');
        }
    }
    
    public static void main(String[] args) 
    		throws Exception {
    	List<String[]> resultOld = ExcelRead.readExcel("d:/tmp/bom_old.xlsx");
    	System.out.println(ExcelRead.formatJson(JSON.toJSONString(resultOld)));
    	
    	System.out.println("--------------------------------");
    	
    	List<String[]> resultNew = ExcelRead.readExcel("d:/tmp/bom_new.xlsx");
    	System.out.println(ExcelRead.formatJson(JSON.toJSONString(resultNew)));
    }
} 