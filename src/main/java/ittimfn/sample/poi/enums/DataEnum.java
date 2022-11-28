package ittimfn.sample.poi.enums;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ittimfn.sample.poi.model.DataModel;

public enum DataEnum {
    SEISU(FormatEnum.LONG) {
        @Override
        public void setCellValue(Cell cell, DataModel model) {
            cell.setCellValue(model.getSeisu());
        }

    },
    SHOUSU(FormatEnum.DOUBLE) {
        @Override
        public void setCellValue(Cell cell, DataModel model) {
            cell.setCellValue(model.getShousu());
        }
    },
    MOJIRETSU(FormatEnum.STRING) {
        @Override
        public void setCellValue(Cell cell, DataModel model) {
            cell.setCellValue(model.getMojiretsu());
        }
    },
    DATE(FormatEnum.DATE) {
        @Override
        public void setCellValue(Cell cell, DataModel model) {
            cell.setCellValue(model.getDate());
        }
    },
    WEEK_OF_DAY(FormatEnum.WEEK_OF_DAY) {
        @Override
        public void setCellValue(Cell cell, DataModel model) {
            cell.setCellValue(model.getDate());
        }
    };

    ;
    
    private FormatEnum format;
    
    private DataEnum(FormatEnum format) {
        this.format = format;
    }
    public void setCellStyle(Cell cell, Workbook workbook) {
        this.format.setCellStyle(cell, workbook);
    }
    public abstract void setCellValue(Cell cell, DataModel model);

}
