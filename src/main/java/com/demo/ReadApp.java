package com.demo;
import com.alibaba.fastjson.JSON;
import java.util.ArrayList;
import java.util.List;

public class ReadApp {
    private static String readExcelPath = ReadApp.class.getClassLoader().getResource("./excelFiles/readExcel.xlsx").getPath();
    public static void main(String[] args) {
        List<JavaBean> javaBeanList = new ArrayList<>(); // 结果
        StringBuilder errorMessage = new StringBuilder();// 记录错误信息
        try {
            ExcelUtil excelUtil = new ExcelUtil();
            excelUtil.initRead(readExcelPath);
            // 上传文件总行数
            int totalRows = excelUtil.getRowCount();
            // 读取索引为0的工作表
            excelUtil.setSheetNum(0);
            for (int i = 1; i <= totalRows; i++) {
                JavaBean javaBean = new JavaBean();
                // 获取一行数据
                int lineNum = i + 1;
                String[] rowData = excelUtil.readExcelLine(i);
                if (null == rowData) {
                    errorMessage.append("第&nbsp" + (lineNum) + "&nbsp行发生错误!\n");
                    continue;
                }
                //获取学号
                int colIndex = 0;
                String id = "";
                try {
                    id = rowData[colIndex];
                } catch (ArrayIndexOutOfBoundsException e1) {
                    errorMessage.append("第&nbsp" + (lineNum) + "&nbsp行数据不全!\n");
                    continue;
                }
                javaBean.setId(id.trim());
                //获取姓名
                colIndex+=1;
                String name = "";
                try {
                    name = rowData[colIndex];
                } catch (ArrayIndexOutOfBoundsException e1) {
                    errorMessage.append("第&nbsp" + (lineNum) + "&nbsp行数据不全!\n");
                    continue;
                }
                javaBean.setName(name.trim());
                //获取性别
                colIndex+=1;
                String sex = "";
                try {
                    sex = rowData[colIndex];
                } catch (ArrayIndexOutOfBoundsException e1) {
                    errorMessage.append("第&nbsp" + (lineNum) + "&nbsp行数据不全!\n");
                    continue;
                }
                javaBean.setSex(sex.trim());
                //获取年龄
                colIndex+=1;
                int age = -1;
                try {
                    age = Integer.parseInt(rowData[colIndex].trim());
                } catch (ArrayIndexOutOfBoundsException e1) {
                    errorMessage.append("第&nbsp" + (lineNum) + "&nbsp行数据不全!\n");
                    continue;
                }
                javaBean.setAge(age);
                javaBeanList.add(javaBean);
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        System.out.println(JSON.toJSONString(javaBeanList));
        System.out.println(errorMessage);
    }
}
