package com.demo;
public class WriteApp {
    public static void main(String[] args) {
        String writeExcelPath = "d:/writeFile.xlsx";
        try {
            ExcelUtil excelUtil = new ExcelUtil();
            excelUtil.initWrite(writeExcelPath);

            excelUtil.setSheetNum(0);
            excelUtil.createRow(0);
            excelUtil.setRowNum(0);
            /* 头部数据项写入 */
            String[] heads = {"学号", "姓名", "性别", "年龄"};
            for (int i = 0; i < heads.length; i++) {
                excelUtil.setCell(i, heads[i]);
            }
            /* 正文数据项写入 */
            String[][] data = {{"123","liming","man","11"},
                    {"456","lihong","woman","11"},
                    {"789","ligang","man","11"}};
            if (data != null && data.length != 0 && data[0].length != 0) {
                //  表格内容的行数
                int rowNumsOfContent = 1;
                for (int i = 0; i < data.length; i++) {
                    int column = 0;
                    excelUtil.createRow(rowNumsOfContent++);
                    excelUtil.setRowNum(rowNumsOfContent);
                    excelUtil.setCell(column++, data[i][0]);
                    excelUtil.setCell(column++, data[i][1]);
                    excelUtil.setCell(column++, data[i][2]);
                    excelUtil.setCell(column++, data[i][3]);
                }
            }
            //表格输出
            excelUtil.export();
        } catch(Exception e){
            e.printStackTrace();
        }
    }
}
