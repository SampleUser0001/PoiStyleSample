package ittimfn.sample.poi.enums;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Getter;

@Getter
public enum FormatEnum {
    LONG("#,##0"),
    DOUBLE("#,##0.0"),
    STRING("");
    
    private String format;
    
    private FormatEnum(String format) {
        this.format = format;
    }
    
    public void setCellStyle(XSSFCell cell, XSSFWorkbook workbook) {
        CellStyle longStyle = workbook.createCellStyle();
        longStyle.setDataFormat(workbook.createDataFormat().getFormat(this.format));
        cell.setCellStyle(longStyle);

    }
}
