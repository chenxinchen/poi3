package cn.chenxinchen.commons.pojo;

import cn.chenxinchen.commons.annotation.ColumnMapping;
import cn.chenxinchen.commons.annotation.ColumnSerial;
import cn.chenxinchen.commons.annotation.RowMapping;
import lombok.Data;

@Data
@RowMapping(2)
public class Goods {
    @ColumnMapping(ColumnSerial.C)
    private int id;
    @ColumnMapping(ColumnSerial.E)
    private String goodsName;
    @ColumnMapping(ColumnSerial.F)
    private double price;
    @ColumnMapping(ColumnSerial.A)
    private Double total;
}
