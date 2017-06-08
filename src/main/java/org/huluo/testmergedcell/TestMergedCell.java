package org.huluo.testmergedcell;
import jxl.Workbook;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 测试下jxl提供的单元格合并功能
 */
public class TestMergedCell {
    public static void main(String[] args) throws IOException, WriteException, BiffException {
        FileOutputStream fileOutputStream = new FileOutputStream("excelResult/xxx.xls");

        WritableWorkbook writableWorkbook = Workbook.createWorkbook(fileOutputStream);

        WritableSheet writableSheet1 = writableWorkbook.createSheet("第一页", 0);

        WritableFont wf_table = new WritableFont(WritableFont.COURIER, 11,
                WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE,
                jxl.format.Colour.BLACK); // 定义格式 字体 下划线 斜体 粗体 颜色
        WritableCellFormat wcf_table2 = new WritableCellFormat(wf_table);
        wcf_table2.setBackground(jxl.format.Colour.YELLOW2);
        wcf_table2.setAlignment(jxl.format.Alignment.CENTRE);
        wcf_table2.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN,jxl.format.Colour.BLACK);

        Label label = new Label(2, 8, "new HelloWorld();",wcf_table2);


        writableSheet1.addCell(label);

        writableSheet1.mergeCells(2, 8, 3, 9);



        writableWorkbook.write();
        writableWorkbook.close();
    }
}
