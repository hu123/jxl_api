package org.huluo.util.exceltemplatefile;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.WritableWorkbook;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.*;

public class StudentScoreExcelTemplateUtilTest {
    @Test
    public void testGenerateExcelTemplate() throws Exception {
        StudentScoreExcelTemplateUtil.generateExcelTemplate("我随机生成的.xls");

    }

    @Test
    public void testReadData() throws Exception {
        Workbook workbook = Workbook.getWorkbook(new File("excelResult/我随机生成的.xls"));
        Sheet sheet = workbook.getSheet(0);
        Cell cell = sheet.getCell(2, 2);

        System.out.println(cell.getContents());



    }
}