package com.view.office.commons.excel;


/**
 * excel单元格合并数据
 */
public class MergedRegions {

    private Integer firstRow;

    private Integer lastRow;

    private Integer firstColumn;

    private Integer lastColumn;

    public MergedRegions() {
    }

    public MergedRegions(Integer firstRow, Integer lastRow, Integer firstColumn, Integer lastColumn) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
    }

    public Integer getFirstRow() {
        return firstRow;
    }

    public void setFirstRow(Integer firstRow) {
        this.firstRow = firstRow;
    }

    public Integer getLastRow() {
        return lastRow;
    }

    public void setLastRow(Integer lastRow) {
        this.lastRow = lastRow;
    }

    public Integer getFirstColumn() {
        return firstColumn;
    }

    public void setFirstColumn(Integer firstColumn) {
        this.firstColumn = firstColumn;
    }

    public Integer getLastColumn() {
        return lastColumn;
    }

    public void setLastColumn(Integer lastColumn) {
        this.lastColumn = lastColumn;
    }
}
