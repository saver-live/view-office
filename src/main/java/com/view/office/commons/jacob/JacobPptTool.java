package com.view.office.commons.jacob;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

public class JacobPptTool {


    public boolean saveAsPdf(String srcFilePath, String pdfFilePath) {
        ComThread.InitMTA(true);
        ActiveXComponent app = null;

        Dispatch ppt = null;
        try {

            app = new ActiveXComponent("PowerPoint.Application");
            Dispatch ppts = app.getProperty("Presentations").toDispatch();

            // 因POWER.EXE的发布规则为同步，所以设置为同步发布
            ppt = Dispatch.call(ppts, "Open", srcFilePath, true,// ReadOnly
                    true,// Untitled指定文件是否有标题
                    false// WithWindow指定文件是否可见
            ).toDispatch();

            Dispatch.call(ppt, "SaveAs", pdfFilePath, 32); //ppSaveAsPDF为特定值32

            return true; // set flag true;
        } catch (ComFailException e) {
            return false;
        } catch (Exception e) {
            return false;
        } finally {
            if (ppt != null) {
                Dispatch.call(ppt, "Close");
            }
            if (app != null) {
                app.invoke("Quit");
            }
            ComThread.Release();
        }
    }
}
