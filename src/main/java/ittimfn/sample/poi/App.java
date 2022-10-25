package ittimfn.sample.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ittimfn.sample.poi.enums.DataEnum;
import ittimfn.sample.poi.enums.FormatEnum;
import ittimfn.sample.poi.model.DataModel;

/**
 * Hello world!
 *
 */
public class App {
    public static final String EXPORT_PATH = Paths.get(System.getProperty("user.dir"), "export", "export.xlsx").toString();

    public void exec(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        try(FileOutputStream fos = new FileOutputStream(EXPORT_PATH);) {
            XSSFRow row = sheet.createRow(0);
            int cellIndex = 0;
            
            DataModel dataModel = new DataModel();
            dataModel.setSeisu(1000000l);
            dataModel.setShousu(111111.11d);
            dataModel.setMojiretsu("hogehgoe");
            dataModel.setDate(new Date());

            // 本当はDataEnum.values()を使用すればループが使えるが、型ごとに確認したいため、分けて記載する。
            // 整数
            XSSFCell cell = row.createCell(cellIndex++);
            DataEnum.SEISU.setCellStyle(cell, workbook);
            DataEnum.SEISU.setCellValue(cell, dataModel);

            // 小数
            cell = row.createCell(cellIndex++);
            DataEnum.SHOUSU.setCellStyle(cell, workbook);
            DataEnum.SHOUSU.setCellValue(cell, dataModel);
            
            // 文字列
            cell = row.createCell(cellIndex++);
            DataEnum.MOJIRETSU.setCellStyle(cell, workbook);
            DataEnum.MOJIRETSU.setCellValue(cell, dataModel);
            
            // 日付
            cell = row.createCell(cellIndex++);
            DataEnum.DATE.setCellStyle(cell, workbook);
            DataEnum.DATE.setCellValue(cell, dataModel);
            
            // 曜日
            cell = row.createCell(cellIndex++);
            DataEnum.WEEK_OF_DAY.setCellStyle(cell, workbook);
            DataEnum.WEEK_OF_DAY.setCellValue(cell, dataModel);
            
            workbook.write(fos);
            fos.close();
            workbook.close();
        }
    }
    
    public void setCellValue() {
        
    }

    public static void main( String[] args ) throws IOException {
        new App().exec(args);
    }

}
