package ittimfn.sample.poi.model;

import java.util.Date;

import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
public class DataModel {
    private long seisu;
    private double shousu;
    private String mojiretsu;
    private Date date;
}
