package com.view.office.commons.excel;

import java.util.List;

/**
 * excel单页文件数据
 */
public class ExcelSheet {

    private String name;
    //excel最大列
    private Integer maxColSize;

    private Integer maxRowSize;

    private List<Cell> data;

    private List<ColumnWidth> widths;

    private List<RowHeight> heights;

    private List<MergedRegions> mergedRegions;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ExcelSheet() {
        this.maxColSize = 0;
    }

    public Integer getMaxColSize() {
        return maxColSize;
    }

    public void setMaxColSize(Integer maxColSize) {
        this.maxColSize = maxColSize;
    }

    public List<Cell> getData() {
        return data;
    }

    public void setData(List<Cell> data) {
        this.data = data;
    }

    public List<MergedRegions> getMergedRegions() {
        return mergedRegions;
    }

    public void setMergedRegions(List<MergedRegions> mergedRegions) {
        this.mergedRegions = mergedRegions;
    }

    public Integer getMaxRowSize() {
        return maxRowSize;
    }

    public void setMaxRowSize(Integer maxRowSize) {
        this.maxRowSize = maxRowSize;
    }

    public List<ColumnWidth> getWidths() {
        return widths;
    }

    public void setWidths(List<ColumnWidth> widths) {
        this.widths = widths;
    }

    public List<RowHeight> getHeights() {
        return heights;
    }

    public void setHeights(List<RowHeight> heights) {
        this.heights = heights;
    }
}
