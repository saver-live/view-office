package com.view.office.commons.excel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class ExcelUtil {

    public List<ExcelSheet> getExcelData(String path) throws IOException {
        Integer col = 0;
        File file = new File(path);
        FileInputStream inputStream = new FileInputStream(file);
        String suffix = path.substring(path.lastIndexOf(".") + 1);
        List<ExcelSheet> excel = new ArrayList<>();
        // 比较文件扩展名
        if (suffix.equalsIgnoreCase("xls")) {
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(inputStream);
            int size = hssfWorkbook.getNumberOfSheets();
            for (int numSheet = 0; numSheet < size; numSheet++) {
                HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
                if (hssfSheet == null) {
                    continue;
                }
                List<CellRangeAddress> mergedRegions = hssfSheet.getMergedRegions();
                List<MergedRegions> regions = new ArrayList<>();
                ExcelSheet sheet = new ExcelSheet();

                setRegions(sheet, mergedRegions, regions);
                sheet.setName(hssfSheet.getSheetName());
                List<Cell> list = new ArrayList<>();
                // 处理当前页，循环读取每一行
                sheet.setMaxRowSize(hssfSheet.getLastRowNum());
                // 遍历改行，获取处理每个cell元素
                List<ColumnWidth> lcw = new ArrayList<>();
                List<RowHeight> lrh = new ArrayList<>();
                for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {

                    // HSSFRow表示行
                    HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                    if (hssfRow == null) {
                        continue;
                    }
                    int minColIx = hssfRow.getFirstCellNum();
                    int maxColIx = hssfRow.getLastCellNum();
                    if (maxColIx > col) {
                        col = maxColIx;
                    }
                    RowHeight rh = new RowHeight();
                    float heightInPoints = hssfRow.getHeightInPoints();
                    rh.setHeight(heightInPoints);
                    rh.setRowIndex(rowNum);
                    lrh.add(rh);
                    for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                        // HSSFCell 表示单元格

                        if (rowNum == 0) {
                            ColumnWidth cw = new ColumnWidth();
                            float columnWidthInPixels = hssfSheet.getColumnWidthInPixels(colIx);
                            cw.setColumnIndex(colIx);
                            cw.setWidth(columnWidthInPixels);
                            lcw.add(cw);
                        }

                        HSSFCell cell = hssfRow.getCell(colIx);
                        if (cell == null) {
                            continue;
                        }
                        HSSFCellStyle cellStyle = cell.getCellStyle();
                        Cell customCell = new Cell();

                        short alignment = cellStyle.getAlignment();
                        customCell.setAlign(alignment);
                        HSSFColor fillForegroundColorColor = cellStyle.getFillForegroundColorColor();
                        if (fillForegroundColorColor != null) {
                            String hexString = fillForegroundColorColor.getHexString();
                            StringBuffer background = new StringBuffer("#");

                            String[] split = hexString.split(":");
                            for (int i = 0; i < split.length; i++) {

                                if (split[i].length() == 4) {
                                    background.append(split[i].substring(2));
                                }
                                if (split[i].length() == 1) {
                                    background.append("0" + split[i]);
                                }
                            }
                            String s = background.toString();
                            if (s.equals("#000000")) {
                                s = "#FFFFFF";
                            }
                            customCell.setBackground(s);
                        }
                        customCell.setBorderBottom(cellStyle.getBorderBottom());
                        customCell.setBorderTop(cellStyle.getBorderTop());
                        customCell.setBorderLeft(cellStyle.getBorderLeft());
                        customCell.setBorderRight(cellStyle.getBorderRight());

                        String stringVal = getStringVal(cell);
                        if (stringVal.contains("\n") || stringVal.contains("\r\n")) {
                            stringVal = stringVal.replace("\n", "<br>");
                            stringVal = stringVal.replace("\r\n", "<br>");
                        }
                        customCell.setText(stringVal);
                        customCell.setY(rowNum);
                        customCell.setX(colIx);
                        list.add(customCell);
                    }
                    sheet.setWidths(lcw);
                    sheet.setHeights(lrh);
                    sheet.setMaxColSize(col);
                    sheet.setData(list);
                }
                excel.add(sheet);
            }
        } else if (suffix.equalsIgnoreCase("xlsx")) {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream);
            // 循环每一页，并处理当前循环页
            Integer size = xssfWorkbook.getNumberOfSheets();
            for (int i = 0; i < size; i++) {
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(i);
                if (xssfSheet == null) {
                    continue;
                }

                List<CellRangeAddress> mergedRegions = xssfSheet.getMergedRegions();
                List<MergedRegions> regions = new ArrayList<>();

                ExcelSheet sheet = new ExcelSheet();

                setRegions(sheet, mergedRegions, regions);
                sheet.setName(xssfSheet.getSheetName());
                List<Cell> list = new ArrayList<>();
                sheet.setMaxRowSize(xssfSheet.getLastRowNum());
                // 处理当前页，循环读取每一行
                List<ColumnWidth> lcw = new ArrayList<>();
                List<RowHeight> lrh = new ArrayList<>();
                for (int rowNum = 0; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                    XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                    if (xssfRow == null) {
                        continue;
                    }
                    int minColIx = xssfRow.getFirstCellNum();
                    int maxColIx = xssfRow.getLastCellNum();
                    if (maxColIx > col) {
                        col = maxColIx;
                    }
                    //行高
                    RowHeight rh = new RowHeight();
                    float heightInPoints = xssfRow.getHeightInPoints();
                    rh.setHeight(heightInPoints);
                    rh.setRowIndex(rowNum);
                    lrh.add(rh);
                    for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                        if (rowNum == 0) {
                            //列宽
                            ColumnWidth cw = new ColumnWidth();
                            float columnWidth = xssfSheet.getColumnWidthInPixels(colIx);
                            cw.setColumnIndex(colIx);
                            cw.setWidth(columnWidth);
                            lcw.add(cw);
                        }

                        XSSFCell cell = xssfRow.getCell(colIx);
                        if (cell == null) {
                            continue;
                        }
                        cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                        XSSFRichTextString rich = cell.getRichStringCellValue();
                        String stringVal;
                        StringBuffer contents = new StringBuffer();
                        if (rich.hasFormatting()) {
                            int startIndex = 0;
                            for (CTRElt ct : rich.getCTRst().getRList()) {
                                XSSFFont font = rich.getFontAtIndex(startIndex);
                                startIndex += ct.getT().length();
                                if (ct.getT().contains("\n") || ct.getT().contains("\r\n")) {
                                    contents.append("<br>");
                                }
                                contents.append("<font style=\"" + fontStyleDetail(font) + " \">" + ct.getT() + "</font>");
                            }
                            stringVal = contents.toString().replace("£", "<font class='font30'>£</font>");
                        } else {
                            stringVal = rich.getString();
                            if (stringVal.contains("\n") || stringVal.contains("\r\n")) {
                                stringVal = stringVal.replace("\n", "<br>");
                                stringVal = stringVal.replace("\r\n", "<br>");
                            }
                        }

                        XSSFCellStyle cellStyle = cell.getCellStyle();
                        XSSFColor fillForegroundXSSFColor = cellStyle.getFillForegroundXSSFColor();

                        Cell customCell = new Cell();
                        if (fillForegroundXSSFColor != null) {
                            String argbHex = fillForegroundXSSFColor.getARGBHex();
                            customCell.setBackground("#" + argbHex.substring(2));
                        }
                        short alignment = cellStyle.getAlignment();
                        customCell.setAlign(alignment);

                        XSSFFont font = cellStyle.getFont();

                        boolean bold = font.getBold();
                        customCell.setBold((short) (bold ? 1 : 0));
                        short color = font.getColor();
                        customCell.setColor(color);
                        short fontHeightInPoints = (short) (font.getFontHeightInPoints());
                        customCell.setFontHeightInPoints(fontHeightInPoints);

                        customCell.setBorderRightColor(cellStyle.getRightBorderColor());
                        customCell.setBorderLeftColor(cellStyle.getLeftBorderColor());
                        customCell.setBorderTopColor(cellStyle.getTopBorderColor());
                        customCell.setBorderBottomColor(cellStyle.getBottomBorderColor());

                        customCell.setBorderBottom(cellStyle.getBorderBottom());
                        customCell.setBorderTop(cellStyle.getBorderTop());
                        customCell.setBorderLeft(cellStyle.getBorderLeft());
                        customCell.setBorderRight(cellStyle.getBorderRight());
                        customCell.setX(colIx);
                        customCell.setY(rowNum);
                        customCell.setText(stringVal);
                        list.add(customCell);
                    }
                    sheet.setWidths(lcw);
                    sheet.setHeights(lrh);
                    sheet.setMaxColSize(col);
                    sheet.setData(list);
                }
                excel.add(sheet);
            }
        }
        inputStream.close();
        return excel;
    }


    /**
     * 获取表格中数据
     *
     * @param cell 表格格子
     * @return
     */
    public static String getStringVal(HSSFCell cell) {
        int cellType = cell.getCellType();
        String value = "";
        switch (cellType) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                double numericCellValue = cell.getNumericCellValue();
                value = String.valueOf(numericCellValue);
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                break;
            case HSSFCell.CELL_TYPE_FORMULA:
                value = cell.getCellFormula();
                break;
            case HSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
        }
        return value;
    }

    public void setRegions(ExcelSheet sheet, List<CellRangeAddress> mergedRegions, List<MergedRegions> regions) {
        for (int j = 0; j < mergedRegions.size(); j++) {
            CellRangeAddress address = mergedRegions.get(j);
            MergedRegions regions1 = new MergedRegions();
            regions1.setFirstColumn(address.getFirstColumn());
            regions1.setLastColumn(address.getLastColumn());
            regions1.setFirstRow(address.getFirstRow());
            regions1.setLastRow(address.getLastRow());
            regions.add(regions1);
        }
        sheet.setMergedRegions(regions);
    }

    private String fontStyleDetail(Font font) {
        StringBuffer buf = new StringBuffer();
        XSSFFont font1 = (XSSFFont) font;
        if (!(font1.getXSSFColor() == null || font1.getXSSFColor().isAuto()))
            buf.append("color:#" + font1.getXSSFColor().getARGBHex().substring(2) + ";");
        if (font.getBold()) {
            buf.append("font-weight: bold;");
        } else {
            buf.append("font-weight: normal;");
        }
        if (font.getItalic())
            buf.append("font-style:italic;");
        buf.append("font-family:" + font.getFontName() + ";");
        buf.append("font-size:" + font.getFontHeightInPoints() + "pt;");
        return buf.toString();
    }
}