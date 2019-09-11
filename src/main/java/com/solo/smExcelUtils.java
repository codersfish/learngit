package com.solo;

import com.google.gson.Gson;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;
/*
 * 1.读取Excel: ReadExcel(String path,int sheetIndex)
 * 2.导出Excel: CreateExcel(List<Map<String,String>> list,String destPath)
 * 3.导出Excel  CreateExcel(String jsonstr,String destPath)
 * Author: ANDY
 * Date: 2019-08-28
 */
public class smExcelUtils {
    //读取Excel
    public static List<Map<String,String>> ReadExcel(String path,int sheetIndex)throws Exception{
        String[] arr=path.split("\\.");
        String suffix=arr[arr.length-1];
        if(suffix.equals("xlsx")){
            return ReadXLSX(path,sheetIndex);
        }
        else if(suffix.equals("xls")){
            return ReadXLS(path,sheetIndex);
        }
        else
            return null;
    }
    //读取xls
    private static List<Map<String,String>> ReadXLS(String path,int sheetIndex) throws Exception{
        List<Map<String,String>> list=new ArrayList<>();
        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);
        HSSFWorkbook book = new HSSFWorkbook(fis);
        HSSFSheet sheet = book.getSheetAt(0);
        int rows = sheet.getPhysicalNumberOfRows();
        int cols=sheet.getRow(0).getPhysicalNumberOfCells();
        for(int i=1;i<rows;i++){
            Map<String,String> map=new LinkedHashMap();
            for(int j=0;j<cols;j++){
                HSSFCell cell=sheet.getRow(i).getCell(j);
                map.put(sheet.getRow(0).getCell(j).toString(),getXlsCellFormatValue(cell));
            }
            list.add(map);
        }
        return list;
    }
    //读取xlsx
    private static List<Map<String,String>> ReadXLSX(String path,int sheetIndex)throws Exception{
        List<Map<String,String>> list=new ArrayList<>();
        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook book = new XSSFWorkbook(fis);
        XSSFSheet sheet = book.getSheetAt(0);
        int rows = sheet.getPhysicalNumberOfRows();
        int cols=sheet.getRow(0).getPhysicalNumberOfCells();
        for(int i=1;i<rows;i++){
            Map map=new LinkedHashMap();
            for(int j=0;j<cols;j++){
                XSSFCell cell=sheet.getRow(i).getCell(j);
                map.put(sheet.getRow(0).getCell(j).toString(),getXlsxCellFormatValue(cell));
            }
            list.add(map);
        }
        return list;
    }
    //获取xls单元格内容
    private static String getXlsCellFormatValue(HSSFCell cell) throws Exception{

        String cellvalue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case NUMERIC:
                    if(HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);
                    }else{
                        HSSFDataFormatter df1=new HSSFDataFormatter();
                        cellvalue = String.valueOf(df1.formatCellValue(cell));
                    }
                    break;
                case FORMULA: {
                    cellvalue=String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case STRING:
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                default:
                    cellvalue = " ";
            }
        }
        else {
            cellvalue = "";
        }

        return cellvalue;
    }
    //获取xlsx单元格内容
    private static String getXlsxCellFormatValue(XSSFCell cell) throws Exception{

        String cellvalue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case NUMERIC:
                    if(HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);
                    }else{
                        HSSFDataFormatter df1=new HSSFDataFormatter();
                        cellvalue = String.valueOf(df1.formatCellValue(cell));
                    }
                    break;
                case FORMULA: {
                    cellvalue=String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case STRING:
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                default:
                    cellvalue = " ";
            }
        }
        else {
            cellvalue = "";
        }

        return cellvalue;
    }
    //创建excel(传入json)
    public static void CreateExcel(String jsonstr,String destPath)throws Exception{
        Gson gson=new Gson();
        Object obj = gson.fromJson(jsonstr, Object.class);
        List<Map<String,String>> list=(List<Map<String,String>>)obj;
        CreateExcel(list,destPath);
    }
    //创建excel(传入List<Map<String,String>>)
    public static void CreateExcel(List<Map<String,String>> list,String destPath) throws Exception{
        // 1. 创建一个工作薄
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建一个目录和文件名
        FileOutputStream out = new FileOutputStream(new File(destPath));
        // 2. 创建一个工作表
        XSSFSheet spreadsheet = workbook.createSheet("Sheet1");

        XSSFCellStyle headerStyle=HeaderColorSytle(workbook);
        XSSFCellStyle int_odd=INT_OddRowColorSytle(workbook);
        XSSFCellStyle dec_odd=DEC_OddRowColorSytle(workbook);
        XSSFCellStyle def_odd=DEF_OddRowColorSytle(workbook);
        XSSFCellStyle int_even=INT_EvenRowColorSytle(workbook);
        XSSFCellStyle dec_even=DEC_EvenRowColorSytle(workbook);
        XSSFCellStyle def_even=DEF_EvenRowColorSytle(workbook);

        //获取列名
        Map map=list.get(0);
        List<String> cols=new ArrayList<>();
        for(Object key : map.keySet()){
            cols.add(key.toString());
        }
        //设置列名(第一)行
        XSSFRow row = spreadsheet.createRow(0);
        for(int i=0;i<cols.size();i++){
            XSSFCell cell = row.createCell(i);
            // 设置单元格的值
            cell.setCellValue(cols.get(i));
            cell.setCellStyle(headerStyle);
        }
        //设置数据行
        for(int i=0;i<list.size();i++){
            XSSFRow datarow = spreadsheet.createRow(i+1);
            Map datamap=list.get(i);
            for(int j=0;j<cols.size();j++){
                XSSFCell cell = datarow.createCell(j);
                String val = datamap.get(cols.get(j)).toString();
                //根据数据类型及奇偶数行,分别设置样式
                if(isMinusNumeric(val).equals("DEC")){//小数
                    if((i+1)%2==0)
                        cell.setCellStyle(dec_even);
                    else
                        cell.setCellStyle(dec_odd);
                    cell.setCellValue(Double.parseDouble(val));
                }
                else if(isMinusNumeric(val).equals("INT")){//整数
                    XSSFCellStyle cellStyle = workbook.createCellStyle();
                    if((i+1)%2==0)
                        cell.setCellStyle(int_even);
                    else
                        cell.setCellStyle(int_odd);
                    cell.setCellValue(Long.parseLong(val));//这里需要注意数据长度
                }
                else{//其它数据类型
                    if((i+1)%2==0)
                        cell.setCellStyle(def_even);
                    else
                        cell.setCellStyle(def_odd);
                    cell.setCellValue(val);
                }
            }
        }
        //设置Excel的自动筛选
        CellRangeAddress c = new CellRangeAddress(0, 0, 0, cols.size()-1);
        spreadsheet.setAutoFilter(c);
        setSizeColumn(spreadsheet);
        workbook.write(out);
        out.close();
    }
    //单元格基本样式(边框,字体)
    private static XSSFCellStyle BasicStyle(XSSFWorkbook workbook){
        // 单元格样式 试验
        XSSFCellStyle basicStyle = workbook.createCellStyle();
        basicStyle.setAlignment(HorizontalAlignment.CENTER);
        basicStyle.setBorderBottom(BorderStyle.THIN);
        basicStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        basicStyle.setBorderLeft(BorderStyle.THIN);
        basicStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        basicStyle.setBorderRight(BorderStyle.THIN);
        basicStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        basicStyle.setBorderTop(BorderStyle.THIN);
        basicStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        Font font=workbook.createFont();
        font.setFontName("Microsoft YaHei");
        font.setColor(IndexedColors.BLACK.getIndex());
        basicStyle.setFont(font);
        return basicStyle;
    }
    //表头样式
    private static XSSFCellStyle HeaderColorSytle(XSSFWorkbook workbook){
        //XSSFColor grey = new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap());
        XSSFCellStyle headerColorStyle = workbook.createCellStyle();
        headerColorStyle.cloneStyleFrom(BasicStyle(workbook));
        headerColorStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        headerColorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font=workbook.createFont();
        font.setFontName("Microsoft YaHei");
        font.setColor(IndexedColors.WHITE.getIndex());
        headerColorStyle.setFont(font);
        return headerColorStyle;
    }
    //奇数行样式
    private static XSSFCellStyle INT_OddRowColorSytle(XSSFWorkbook workbook){
        //XSSFColor grey = new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap());
        XSSFCellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(BasicStyle(workbook));
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font=workbook.createFont();
        font.setFontName("Microsoft YaHei");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0"));
        style.setFont(font);
        return style;
    }
    private static XSSFCellStyle DEC_OddRowColorSytle(XSSFWorkbook workbook){
        //XSSFColor grey = new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap());
        XSSFCellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(BasicStyle(workbook));
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font=workbook.createFont();
        font.setFontName("Microsoft YaHei");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        style.setFont(font);
        return style;
    }
    private static XSSFCellStyle DEF_OddRowColorSytle(XSSFWorkbook workbook){
        //XSSFColor grey = new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap());
        XSSFCellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(BasicStyle(workbook));
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font=workbook.createFont();
        font.setFontName("Microsoft YaHei");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        return style;
    }
    //偶数行样式
    private static XSSFCellStyle INT_EvenRowColorSytle(XSSFWorkbook workbook){
        //XSSFColor grey = new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap());
        XSSFCellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(BasicStyle(workbook));
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font=workbook.createFont();
        font.setFontName("Microsoft YaHei");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0"));
        return style;
    }
    private static XSSFCellStyle DEC_EvenRowColorSytle(XSSFWorkbook workbook){
        //XSSFColor grey = new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap());
        XSSFCellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(BasicStyle(workbook));
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font=workbook.createFont();
        font.setFontName("Microsoft YaHei");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        return style;
    }
    private static XSSFCellStyle DEF_EvenRowColorSytle(XSSFWorkbook workbook){
        //XSSFColor grey = new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap());
        XSSFCellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(BasicStyle(workbook));
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font=workbook.createFont();
        font.setFontName("Microsoft YaHei");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        return style;
    }
    //判断是否为整数或者小数,整数返回"INT",小数返回"DEC",其余返回"NAN"
    private static String isNumeric(String str){
        Pattern pattern = Pattern.compile("[0-9]*");
        if(str.indexOf(".")>0){//判断是否有小数点
            if(str.indexOf(".")==str.lastIndexOf(".") && str.split("\\.").length==2){ //判断是否只有一个小数点
                if(pattern.matcher(str.replace(".","")).matches())
                    return "DEC";
                else
                    return "NAN";
            }else {
                return "NAN";
            }
        }else {
            if(pattern.matcher(str).matches())
                return "INT";
            else
                return "NAN";
        }
    }
    //这个才是最终被调用的,加上了对负数的处理
    private static String isMinusNumeric(String str){
        String first=str.substring(0,1);
        String result="";
        if(first.equals("-")){
            str=str.substring(1);
        }
        result=isNumeric(str);
        return result;
    }
    // 自适应宽度(中文支持)
    private static void setSizeColumn(XSSFSheet sheet) {
        for (int columnNum = 0; columnNum <= 8; columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                XSSFRow currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }

                if (currentRow.getCell(columnNum) != null) {
                    XSSFCell currentCell = currentRow.getCell(columnNum);
                    if (currentCell.getCellType() == CellType.STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256+1000);
        }
    }
}