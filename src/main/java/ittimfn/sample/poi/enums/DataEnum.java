package ittimfn.sample.poi.enums;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ittimfn.sample.poi.model.DataModel;

public enum DataEnum {
    SEISU(FormatEnum.LONG) {
        @Override
        public void setCellValue(XSSFCell cell, DataModel model) {
            cell.setCellValue(model.getSeisu());
        }

    },
    SHOUSU(FormatEnum.DOUBLE) {
        @Override
        public void setCellValue(XSSFCell cell, DataModel model) {
            cell.setCellValue(model.getShousu());
        }
    },
    MOJIRETSU(FormatEnum.STRING) {
        @Override
        public void setCellValue(XSSFCell cell, DataModel model) {
            cell.setCellValue(model.getMojiretsu());
        }
    };
    
    private FormatEnum format;
    
    private DataEnum(FormatEnum format) {
        this.format = format;
    }
    public void setCellStyle(XSSFCell cell, XSSFWorkbook workbook) {
        this.format.setCellStyle(cell, workbook);
    }
    public abstract void setCellValue(XSSFCell cell, DataModel model);

}
