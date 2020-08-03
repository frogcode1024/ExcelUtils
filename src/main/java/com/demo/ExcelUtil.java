package com.demo;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 * 实现excel读写的工具类，xls、xlsx格式通用
 */
public class ExcelUtil {
    private static final Logger LOG = Logger.getLogger(ExcelUtil.class);
    // 定制浮点数格式
    private static String NUMBER_FORMAT = "#,##0";
    // 定制日期格式
    private static String DATE_FORMAT = "m/d/yy h:mm";
    // 用于读的workbook
    private Workbook workbookRead = null;
    // 用于写的workbook
    private Workbook workbookWrite = null;

    private Sheet sheet = null;

    private Row row = null;

    private int sheetNum = 0;

    private int rowNum = 0;

    private InputStream in = null;

    private OutputStream out = null;

    public ExcelUtil() {
    }

    /**
     * 初始化 read
     */
    public void initRead(String readPath) throws Exception{
        try {
            String fileType = readPath.substring(readPath.lastIndexOf(".")+1);
            File file = new File(readPath);
            in = new FileInputStream(file);
            //创建文档对象
            if (fileType.equals("xls")) {
                workbookRead = new HSSFWorkbook(in);
            } else if (fileType.equals("xlsx")) {
                workbookRead = new XSSFWorkbook(in);
            } else {
                LOG.error("文件后缀格式不正确，只能xls或xlsx!");
                throw new IOException("文件后缀格式不正确，只能xls或xlsx!");
            }
        } catch (Exception e) {
            LOG.error("初始化读取出错");
            throw new Exception("初始化读取出错", e);
        }
    }

    /**
     * 初始化write
     */
    public void initWrite(String writePath) throws Exception {
        try {
            String fileType = writePath.substring(writePath.lastIndexOf(".")+1);
            File file = new File(writePath);
            OutputStream outputStream = new FileOutputStream(file);
            if (fileType.equals("xls")) {
                this.workbookWrite = new HSSFWorkbook();
            } else if (fileType.equals("xlsx")) {
                this.workbookWrite = new XSSFWorkbook();
            } else {
                LOG.error("文件后缀格式不正确，只能xls或xlsx!");
                throw new IOException("文件后缀格式不正确，只能xls或xlsx!");
            }
            this.out = outputStream;
            this.sheet = workbookWrite.createSheet();
        }catch (Exception e){
            LOG.error("初始化写入出错");
            throw new Exception("初始化写入出错", e);
        }
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public void setSheetNum(int sheetNum) {
        this.sheetNum = sheetNum;
    }

    /**
     * sheet表数目
     */
    public int getSheetCount() {
        int sheetCount = -1;
        sheetCount = workbookRead.getNumberOfSheets();
        return sheetCount;
    }

    /**
     * sheetNum下的记录行数
     */
    public int getRowCount() {
        int rowCount = -1;
        if (workbookRead == null) {
            LOG.error("ExcelUtil.getRowCount WorkBook为空!");
        } else {
            Sheet sheet = workbookRead.getSheetAt(this.sheetNum);
            //sheet.getLeftCol();
            rowCount = sheet.getLastRowNum();
        }
        return rowCount;
    }

    /**
     * 读取指定sheetNum的记录行数
     */
    public int getRowCount(int sheetNum) {
        Sheet sheet = workbookRead.getSheetAt(sheetNum);
        int rowCount = -1;
        rowCount = sheet.getLastRowNum();
        return rowCount;
    }

    /**
     * 某行的记录列数
     */
    public int getCellCount(int row) {
        int celCount = -1;
        if (workbookRead == null) {
            LOG.error("ExcelUtil.getCellCount WorkBook为空!");
        } else {
            Sheet sheet = workbookRead.getSheetAt(this.sheetNum);
            celCount = sheet.getRow(row).getLastCellNum(); // getPhysicalNumberOfCells 是获取不为空的列个数; getLastCellNum 是获取最后一个不为空的列是第几个
        }
        return celCount;
    }


    /**
     * 得到指定行的内容
     */
    public String[] readExcelLine(int lineNum) {
        return readExcelLine(this.sheetNum, lineNum);
    }

    /**
     * 指定工作表和指定行的内容
     */
    public String[] readExcelLine(int sheetNum, int lineNum) {
        if (sheetNum < 0 || lineNum < 0)
            return null;
        String[] strExcelLine = null;
        try {
            sheet = workbookRead.getSheetAt(sheetNum);
            row = sheet.getRow(lineNum);

            int cellCount = row.getLastCellNum();
            strExcelLine = new String[cellCount + 1];
            for (int i = 0; i < cellCount; i++) {
                strExcelLine[i] = readStringExcelCell(lineNum, i);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return strExcelLine;
    }

    /**
     * 读取某一行中指定列的内容
     */
    public String readStringExcelCell(int cellNum) {
        return readStringExcelCell(this.rowNum, cellNum);
    }


    /**
     * 指定行和指定列的内容
     */
    public String readStringExcelCell(int rowNum, int cellNum) {
        return readStringExcelCell(this.sheetNum, rowNum, cellNum);
    }

    /**
     * 指定工作表、行、列下的内容
     */
    public String readStringExcelCell(int sheetNum, int rowNum, int cellNum) {
        if (sheetNum < 0 || rowNum < 0)
            return "";
        String strExcelCell = "";
        try {
            sheet = workbookRead.getSheetAt(sheetNum);
            row = sheet.getRow(rowNum);

            if (row.getCell((short) cellNum) != null) { // add this condition
                // judge
                switch (row.getCell((short) cellNum).getCellType()) {
                    case Cell.CELL_TYPE_FORMULA:
                        strExcelCell = "FORMULA ";
                        break;
                    case Cell.CELL_TYPE_NUMERIC: {
                        DecimalFormat df = new DecimalFormat("0");
                        strExcelCell = df.format(row.getCell((short) cellNum).getNumericCellValue());
                        // 解析日期
                        if (DateUtil.isCellDateFormatted(row.getCell((short) cellNum))) {
                            Date date = DateUtil.getJavaDate(row.getCell((short) cellNum).getNumericCellValue());
                            strExcelCell = Date2DetailString(date);
                        }
                    }
                    break;
                    case Cell.CELL_TYPE_STRING:
                        strExcelCell = row.getCell((short) cellNum).getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_BLANK:
                        strExcelCell = "";
                        break;
                    default:
                        strExcelCell = "";
                        break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return strExcelCell;
    }


    /**
     * 导出Excel文件
     */
    public void export() throws Exception {
        try {
            workbookWrite.write(out);
            out.flush();
            out.close();
        } catch (FileNotFoundException e) {
            throw new Exception(" 生成Excel文件出错! ", e);
        } catch (IOException e) {
            throw new Exception(" 写入Excel文件出错! ", e);
        } finally {
            out.flush();
            out.close();
        }

    }

    /**
     * 增加一行
     * @param index 行号
     */
    public void createRow(int index) {
        this.row = this.sheet.createRow(index);
    }

    /**
     * 获取单元格的值
     * @param index 列号
     */
    public String getCell(int index) {
        Cell cell = this.row.getCell((short) index);
        String strExcelCell = "";
        if (cell != null) { // add this condition
            // judge
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_FORMULA:
                    strExcelCell = "FORMULA ";
                    break;
                case Cell.CELL_TYPE_NUMERIC: {
                    strExcelCell = String.valueOf(cell.getNumericCellValue());
                }
                break;
                case Cell.CELL_TYPE_STRING:
                    strExcelCell = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BLANK:
                    strExcelCell = "";
                    break;
                default:
                    strExcelCell = "";
                    break;
            }
        }
        return strExcelCell;
    }

    /**
     * 设置单元格
     * @param index 列号
     * @param value 单元格填充值
     */
    public void setCell(int index, int value) {
        Cell cell = this.row.createCell((short) index);
        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
        cell.setCellValue(value);
    }

    /**
     * 设置单元格
     *
     * @param index 列号
     * @param value 单元格填充值
     */
    public void setCell(int index, double value) {
        Cell cell = this.row.createCell((short) index);
        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
        cell.setCellValue(value);
        CellStyle cellStyle = workbookWrite.createCellStyle(); // 建立新的cell样式
        DataFormat format = workbookWrite.createDataFormat();
        cellStyle.setDataFormat(format.getFormat(NUMBER_FORMAT)); // 设置cell样式为定制的浮点数格式
        cell.setCellStyle(cellStyle); // 设置该cell浮点数的显示格式
    }

    /**
     * 设置单元格
     *
     * @param index 列号
     * @param value 单元格填充值
     */
    public void setCell(int index, String value) {
        Cell cell = this.row.createCell((short) index);
        cell.setCellValue(value);
    }

    /**
     * 设置单元格
     *
     * @param index 列号
     * @param value 单元格填充值
     */
    public void setCell(int index, Calendar value) {
        Cell cell = this.row.createCell((short) index);
        //cell.setEncoding(XLS_ENCODING);
        cell.setCellValue(value.getTime());
        CellStyle cellStyle = workbookWrite.createCellStyle(); // 建立新的cell样式
        DataFormat format = workbookWrite.createDataFormat();
        cellStyle.setDataFormat(format.getFormat(NUMBER_FORMAT)); // 设置cell样式为定制的浮点数格式
        cell.setCellStyle(cellStyle); // 设置该cell日期的显示格式
    }

    public String Date2DetailString(Date date) {
        String result = null;
        SimpleDateFormat sdf = null;

        if (date != null) {
            sdf = new SimpleDateFormat("yyyy-MM-dd");
            result = sdf.format(date);
        }
        return result;
    }


    public CellStyle creatCellStyle() {
        CellStyle css = this.workbookWrite.createCellStyle();
        DataFormat format = this.workbookWrite.createDataFormat();
        css.setDataFormat(format.getFormat("@"));
        return css;

    }

    /**
     * 设置单元格(文本格式)
     * @param index 列号
     * @param value 单元格填充值
     */
    public void setCellWithStringStyle(int index, String value,CellStyle css) {
        Cell cell = this.row.createCell((short) index);
        sheet.setDefaultColumnStyle(index, css);
        cell.setCellType(Cell.CELL_TYPE_STRING);
        cell.setCellValue(value);
    }

}
