package org.huluo.servlet;

import org.springframework.util.StreamUtils;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * 用于下载学生成绩录入的模板excel文件
 */
@WebServlet("/download")
public class DownloadStudentScoreExcelTemplateFileServlet extends HttpServlet {

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        this.doPost(req, resp);
    }

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {

        //强制浏览器弹出下载框,并且解决中文文件名显示-----.xls的问题
        resp.setHeader("Content-Disposition", "attachment;filename=" + new String("学生成绩导入模板.xls".getBytes("utf-8"),"ISO-8859-1"));
        FileInputStream fileInputStream = new FileInputStream("E:/git_tmp/getDataFromTemplateExcelFile/excelResult/我随机生成的.xls");
        StreamUtils.copy(fileInputStream, resp.getOutputStream());
    }
}
