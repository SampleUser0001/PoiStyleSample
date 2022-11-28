package ittimfn.sample.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import ittimfn.sample.poi.enums.DataEnum;
import ittimfn.sample.poi.model.DataModel;

/**
 * Hello world!
 *
 */
public class App {
    public static final String EXPORT_PATH = Paths.get(System.getProperty("user.dir"), "export", "export.xlsx").toString();

    private Workbook workbook;

    public void exec(String[] args) throws IOException {
        this.workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        try(FileOutputStream fos = new FileOutputStream(EXPORT_PATH);) {
            Row row = sheet.createRow(0);
            int cellIndex = 0;
            
            DataModel dataModel = new DataModel();
            dataModel.setSeisu(1000000l);
            dataModel.setShousu(111111.11d);
            dataModel.setMojiretsu("hogehgoe");
            dataModel.setDate(new Date());

            // 本当はDataEnum.values()を使用すればループが使えるが、型ごとに確認したいため、分けて記載する。
            // 整数
            Cell cell = row.createCell(cellIndex++);
            DataEnum.SEISU.setCellStyle(cell, workbook);
            DataEnum.SEISU.setCellValue(cell, dataModel);
            cell.setCellStyle(this.getCellStyle());
            
            // 小数
            cell = row.createCell(cellIndex++);
            DataEnum.SHOUSU.setCellStyle(cell, workbook);
            DataEnum.SHOUSU.setCellValue(cell, dataModel);
            cell.setCellStyle(this.getCellStyle());
            
            // 文字列
            cell = row.createCell(cellIndex++);
            DataEnum.MOJIRETSU.setCellStyle(cell, workbook);
            DataEnum.MOJIRETSU.setCellValue(cell, dataModel);
            cell.setCellStyle(this.getCellStyle());
            
            // 日付
            cell = row.createCell(cellIndex++);
            DataEnum.DATE.setCellStyle(cell, workbook);
            DataEnum.DATE.setCellValue(cell, dataModel);
            cell.setCellStyle(this.getCellStyle());
            
            // 曜日
            cell = row.createCell(cellIndex++);
            DataEnum.WEEK_OF_DAY.setCellStyle(cell, workbook);
            DataEnum.WEEK_OF_DAY.setCellValue(cell, dataModel);
            cell.setCellStyle(this.getCellStyle());

            // プロパティ書き込み
            // 作成者
            // workbook.getProperties().getCoreProperties().setCreator("");

            // // プログラム名
            // workbook.getProperties().getExtendedProperties().getUnderlyingProperties().setApplication("");

            workbook.write(fos);
            fos.close();
            workbook.close();
        }
    }
    
    public CellStyle getCellStyle() {

        // 毎回CellStyleインスタンスを生成するのは本来はNG。Poiで指定できるCellStyleの数は上限がある。
        CellStyle cellStyle = this.workbook.createCellStyle();

        // 色を付ける
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());        
        return cellStyle;
    }


    public static void main( String[] args ) throws IOException {
        new App().exec(args);
    }

}
