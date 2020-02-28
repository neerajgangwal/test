
package com.exceljava.jinx.examples;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import java.io.File;

/**
 *
 * @author neeraj
 */
public class Jacob {
    
    public static void main(String args[]){
        final File file = new File( "D:\\test3.xlsm");
            final String macroName = "DeleteAllEntries";
            callExcelMacro(file, macroName);
    }
    
    private static void callExcelMacro(File file, String macroName) {
            ComThread.InitSTA();

            final ActiveXComponent excel = new ActiveXComponent("Excel.Application");

            try {
                //  This will open the excel if the property is set to true   
                //  Excel.setProperty("Visible", new Variant(true));
            final com.jacob.com.Dispatch workbooks = excel.getProperty("Workbooks") .toDispatch();
            final com.jacob.com.Dispatch workBook = Dispatch.call(workbooks, "Open", file.getAbsolutePath()).toDispatch();
                                        
            //  Calls the macro
            //  final Variant result = Dispatch.call(excel, "Run", new Variant(file.getName()+ macroName));
              final Variant result = Dispatch.call(excel, "Run", macroName);

                // Saves and closes
                Dispatch.call(workBook, "Save");
                                   
                com.jacob.com.Variant f = new com.jacob.com.Variant(true);
                Dispatch.call(workBook, "Close", f);

                } catch (Exception e) {
                          e.printStackTrace();
                } finally {
                    excel.invoke("Quit", new Variant[0]);
                    ComThread.Release();
                }
    }
    
}
