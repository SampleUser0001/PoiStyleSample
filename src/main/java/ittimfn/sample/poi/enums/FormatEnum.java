package ittimfn.sample.poi.enums;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import lombok.Getter;

@Getter
public enum FormatEnum {
    LONG("#,##0"),
    DOUBLE("#,##0.0"),
    STRING(""),
    DATE("yyyy/mm/dd"),
    WEEK_OF_DAY("aaa");
    
    private String format;
    
    private FormatEnum(String format) {
        this.format = format;
    }
    
    public void setCellStyle(Cell cell, Workbook workbook) {
        CellStyle longStyle = workbook.createCellStyle();
        longStyle.setDataFormat(workbook.createDataFormat().getFormat(this.format));
        cell.setCellStyle(longStyle);

    }
}
