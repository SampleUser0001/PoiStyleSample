package ittimfn.sample.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
    public static final String EXPORT_PATH = Paths.get(System.getProperty("user.dir"), "export", "export.xlsx").toString();

    public static void main( String[] args ) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        try(FileOutputStream fos = new FileOutputStream(EXPORT_PATH);) {
            XSSFRow row = sheet.createRow(0);

            int cellIndex = 0;

            // 整数
            XSSFCell cell = row.createCell(cellIndex++);

            CellStyle longStyle = workbook.createCellStyle();
            short longFormat = workbook.createDataFormat().getFormat("#,##0");
            longStyle.setDataFormat(longFormat);

            cell.setCellStyle(longStyle);
            cell.setCellValue(1000000l);
            
            // 小数
            cell = row.createCell(cellIndex++);

            CellStyle doubleStyle = workbook.createCellStyle();
            short doubleFormat = workbook.createDataFormat().getFormat("#,##0.0");
            doubleStyle.setDataFormat(doubleFormat);

            cell.setCellStyle(doubleStyle);
            cell.setCellValue(111111.11d);
            
            // 文字列
            cell = row.createCell(cellIndex++);
            cell.setCellValue("hogehoge");
            
            workbook.write(fos);
            fos.close();
            workbook.close();
        }
    }
}
