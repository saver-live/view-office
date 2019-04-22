package com.view.office.commons.jacob;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class JacobExcelTool {
    private ActiveXComponent xl = null; //Excel对象(防止打开多个)
    private Dispatch workbooks = null;  //工作簿对象
    private Dispatch workbook = null; //具体工作簿
    private Dispatch sheets = null;// 获得sheets集合对象
    private Dispatch currentSheet = null;// 当前sheet

    public ActiveXComponent getXl() {
        return xl;
    }

    public Dispatch getWorkbooks() {
        return workbooks;
    }

    public Dispatch getWorkbook() {
        return workbook;
    }

    /**
     * 打开excel文件
     *
     * @param visible  是否显示打开
     */
    public JacobExcelTool(boolean visible) {
        if (xl == null || xl.m_pDispatch == 0)
            xl = new ActiveXComponent("Excel.Application"); //Excel对象
        xl.setProperty("Visible", new Variant(visible));//设置是否显示打开excel
        if (workbooks == null || workbooks.m_pDispatch == 0)
            workbooks = xl.getProperty("Workbooks").toDispatch(); //工作簿对象

    }

    public void OpenExcel(String filepath, boolean readonly) {
        try {
            initComponents(); //清空原始变量
            ComThread.InitSTA();
	 /*           if(xl==null)
	                xl = new ActiveXComponent("Excel.Application"); //Excel对象
	            xl.setProperty("Visible", new Variant(visible));//设置是否显示打开excel
	            if(workbooks==null)
	                workbooks = xl.getProperty("Workbooks").toDispatch(); //工作簿对象
	  */
            workbook = Dispatch.invoke(  //打开具体工作簿
                    workbooks,
                    "Open",
                    Dispatch.Method,
                    new Object[]{filepath, new Variant(false),
                            new Variant(readonly)},// 是否以只读方式打开
                    new int[1]).toDispatch();
        } catch (Exception e) {
            e.printStackTrace();
            releaseSource();
        }
    }

    /**
     * 工作簿另存为
     *
     * @param filePath 另存为的路径
     */
    public void SaveAs(String filePath) {
        Dispatch.invoke(workbook, "SaveAs", Dispatch.Method,
                new Object[]{filePath,
                        new Variant(44)}, new int[1]);
    }

    public void SavePdf(String filePath) {
        Dispatch.invoke(workbook, "SaveAs", Dispatch.Method,
                new Object[]{filePath,
                        new Variant(57)}, new int[1]);
    }

    public void savePrn(String filePath) {
        Dispatch.invoke(workbook, "SaveAs", Dispatch.Method,
                new Object[]{filePath,
                        new Variant(36)}, new int[1]);
    }

    public void saveHtml(String filePath){
        Dispatch.invoke(workbook, "SaveAs", Dispatch.Method,
                new Object[]{filePath,
                        new Variant(44)}, new int[1]);
    }

    public void SaveAsExc(String filePath) {
        Dispatch.call(workbook, "SaveAs", filePath);
    }

    /**
     * 关闭excel文档
     *
     * @param f 含义不明 （关闭是否保存？默认false）
     */
    public void CloseExcel(boolean f) {
        try {
            Dispatch.call(workbook, "Save");
            Dispatch.call(workbook, "Close", new Variant(f));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            workbook = null;
//            releaseSource();

        }
    }

    /**
     * 释放资源
     */
    public void releaseSource() {
        if (xl != null) {
            xl.invoke("Quit", new Variant[]{});
            xl = null;
        }
        workbooks = null;
        ComThread.Release();
        System.gc();
    }

    /**
     * 添加新的工作表(sheet)，（添加后为默认为当前激活的工作表）
     */
    public Dispatch addSheet() {
        return Dispatch.get(Dispatch.get(workbook, "sheets").toDispatch(), "add").toDispatch();
    }

    /**
     * 修改当前工作表的名字
     *
     * @param newName
     */
    public void modifyCurrentSheetName(String newName) {
        Dispatch.put(getCurrentSheet(), "name", newName);
    }

    /**
     * 得到当前工作表的名字
     *
     * @return
     */
    public String getCurrentSheetName() {
        return Dispatch.get(getCurrentSheet(), "name").toString();
    }

    /**
     * 得到工作薄的名字
     *
     * @return
     */
    public String getWorkbookName() {
        if (workbook == null)
            return null;
        return Dispatch.get(workbook, "name").toString();
    }

    /**
     * 得到sheets的集合对象
     *
     * @return
     */
    public Dispatch getSheets() {
        if (sheets == null)
            sheets = Dispatch.get(workbook, "sheets").toDispatch();
        return sheets;
    }

    /**
     * 得到当前sheet
     *
     * @return
     */
    public Dispatch getCurrentSheet() {
        currentSheet = Dispatch.get(workbook, "ActiveSheet").toDispatch();
        return currentSheet;
    }

    /**
     * 通过工作表名字得到工作表
     *
     * @param name sheetName
     * @return
     */
    public Dispatch getSheetByName(String name) {
        return Dispatch.invoke(getSheets(), "Item", Dispatch.Get, new Object[]{name}, new int[1]).toDispatch();
    }

    /**
     * 通过工作表索引得到工作表(第一个工作簿index为1)
     *
     * @param index
     * @return sheet对象
     */
    public Dispatch getSheetByIndex(Integer index) {
        return Dispatch.invoke(getSheets(), "Item", Dispatch.Get, new Object[]{index}, new int[1]).toDispatch();
    }

    /**
     * 得到sheet的总数
     *
     * @return
     */
    public int getSheetCount() {
        int count = Dispatch.get(getSheets(), "count").toInt();
        return count;
    }

    /**
     * 调用excel宏
     *
     * @param macroName 宏名
     */
    public void callMacro(String macroName) {
        Dispatch.call(xl, "Run", new Variant(macroName));
    }

    /**
     * 单元格写入值
     *
     * @param sheet    被操作的sheet
     * @param position 单元格位置，如：C1
     * @param type     值的属性 如：value
     * @param value
     */
    public void setValue(Dispatch sheet, String position, String type, Object value) {

        Dispatch cell = Dispatch.invoke(sheet, "Range",
                Dispatch.Get, new Object[]{position}, new int[1])
                .toDispatch();
        Dispatch.put(cell, type, value);
    }

    /**
     * 单元格读取值
     *
     * @param position 单元格位置，如： C1
     * @param sheet
     * @return
     */
    public Variant getValue(String position, Dispatch sheet) {
        Dispatch cell = Dispatch.invoke(sheet, "Range", Dispatch.Get,
                new Object[]{position}, new int[1]).toDispatch();
        Variant value = Dispatch.get(cell, "Value");
        return value;
    }

    private void initComponents() {
        workbook = null;
        currentSheet = null;
        sheets = null;
    }
}

