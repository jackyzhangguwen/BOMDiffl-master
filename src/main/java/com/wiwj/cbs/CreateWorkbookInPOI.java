package com.wiwj.cbs;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 使用POI接口创建EXCEL
 */
public class CreateWorkbookInPOI {

    public static void CreateDiffResultExcel(
    		String targetPath,
    		List<String []> resultOld,
    		List<String []> resultNew) 
    		throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("比较结果");

        int rowIdx = 0;
        int columnCount = 10;
        Row rowTitle = sheet.createRow(rowIdx++);
        for (int i = 0; i < columnCount; i++) {
            String label = "列" + (i + 1);
            Cell cellTitle = rowTitle.createCell(i);
            // 设置样式-颜色  
            XSSFCellStyle style = workbook.createCellStyle();    
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);    
            style.setFillForegroundColor(HSSFColor.BRIGHT_GREEN.index);
            cellTitle.setCellStyle(style);
            cellTitle.setCellValue(label);
        }

        for(int i = 0; i < 100; i++) {
	        Row rowContent = sheet.createRow(rowIdx++);
	        for (int j = 0; j < columnCount; j++) {
	            String content = "值" + i + "行" + j + "列";
	            rowContent.createCell(j).setCellValue(content);
	        }
        }
        
        OutputStream outputStream = new FileOutputStream(targetPath + "bom_diff_result.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
 
    }
}