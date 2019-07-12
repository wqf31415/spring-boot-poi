package xyz.wqf31415.controller.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;
import xyz.wqf31415.annotation.ExcelExportField;
import java.lang.reflect.Field;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Created by Administrator on 2019/7/11.
 *
 * @author WeiQuanfu
 */
public class ExcelExportUtil {

    private static final String titleFontName = "宋体";

    private static final short titleFontSize = 13;

    private static final String contentFontName = "宋体";

    private static final short contentFontSize = 11;

    /**
     * 初始化
     * 创建指定的sheet，并在 sheet 中生成标题、表头
     *
     * @return
     */
    public static SXSSFWorkbook initSheetAndTable(SXSSFWorkbook workbook,List<Field> targetFeildList, String sheetName, String tableTitle){
        if (workbook.getSheet(sheetName) != null){
            throw new RuntimeException("sheet named '"+sheetName +"' is exist");
        }
        SXSSFSheet sheet = workbook.createSheet(sheetName);
        CellStyle titleStyle = createTitleStyle(workbook);

        // 表题
        SXSSFRow row0 = sheet.createRow(0);
        SXSSFCell cell_00 = row0.createCell(0, CellType.STRING);
        cell_00.setCellStyle(titleStyle);
        cell_00.setCellValue(tableTitle);

        // 表头
        SXSSFRow row1 = sheet.createRow(1);
        for (int i = 0; i<targetFeildList.size();i++) {
            SXSSFCell cell_0_i = row1.createCell(i, CellType.STRING);
            cell_0_i.setCellStyle(titleStyle);
            Field f = targetFeildList.get(i);
            ExcelExportField anno = f.getAnnotation(ExcelExportField.class);
            String title = anno==null || anno.title() == null || "".equals(anno.title()) ? f.getName():anno.title();
            cell_0_i.setCellValue(title);
        }
        // 冻结前两行
        sheet.createFreezePane(0, 2, 0, 2);

        // tab 颜色
        sheet.setTabColor(new XSSFColor(new java.awt.Color(180,198,231)));

        // 合并第一行，用来放表名
        CellRangeAddress tableTitleRange = new CellRangeAddress(0, 0, 0, targetFeildList.size() - 1);
        for (int i = tableTitleRange.getFirstRow();i<=tableTitleRange.getLastRow();i++){
            SXSSFRow row_i = sheet.getRow(i);
            for (int j = tableTitleRange.getFirstColumn();j<=tableTitleRange.getLastColumn();j++){
                row_i.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellStyle(titleStyle);
            }
        }
        sheet.addMergedRegion(tableTitleRange);


        // 给第二行表头添加过滤选项
        sheet.setAutoFilter(new CellRangeAddress(1,1,0,targetFeildList.size()-1));
        return workbook;
    }

    public static CellStyle createTitleStyle(Workbook workbook){
        CellStyle titleStyle = workbook.createCellStyle();

        // 对齐方式
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);

        // 字体
        Font font = workbook.createFont();
        font.setFontHeightInPoints(titleFontSize);
        font.setFontName(titleFontName);
        titleStyle.setFont(font);

        // 边框
        titleStyle.setBorderLeft(BorderStyle.MEDIUM);
        titleStyle.setBorderRight(BorderStyle.MEDIUM);
        titleStyle.setBorderBottom(BorderStyle.MEDIUM);
        titleStyle.setBorderTop(BorderStyle.MEDIUM);

        // 按内容调整表格宽度
        titleStyle.setShrinkToFit(true);

        return titleStyle;
    }

    public static CellStyle createContentCellStyle(Workbook workbook){
        CellStyle contentCellStyle = workbook.createCellStyle();

        // 对齐方式
//        contentCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//        contentCellStyle.setAlignment(HorizontalAlignment.CENTER);

        // 字体
        Font font = workbook.createFont();
        font.setFontHeightInPoints(contentFontSize);
        font.setFontName(contentFontName);
        contentCellStyle.setFont(font);

        // 边框
        contentCellStyle.setBorderLeft(BorderStyle.THIN);
        contentCellStyle.setBorderRight(BorderStyle.THIN);
        contentCellStyle.setBorderBottom(BorderStyle.THIN);
        contentCellStyle.setBorderTop(BorderStyle.THIN);

        // 按内容调整表格宽度
        contentCellStyle.setShrinkToFit(true);

        return contentCellStyle;
    }

    /**
     * 写入数据到工作簿
     *
     * @param dataList
     * @param <T>
     * @return
     */
    public static <T> SXSSFWorkbook whiteDate(SXSSFWorkbook workbook, List<T> dataList, List<Field> fieldList, String sheetName){
        if (workbook == null || CollectionUtils.isEmpty(dataList) || CollectionUtils.isEmpty(fieldList) || StringUtils.isEmpty(sheetName)) return null;
        SXSSFSheet sheet = workbook.getSheet(sheetName);
        CellStyle contentCellStyle = createContentCellStyle(workbook);
        for (int i = 0 ;i<dataList.size();i++){
            SXSSFRow row_i = sheet.createRow(i+2);
            T data = dataList.get(i);
            for (int j = 0; j<fieldList.size();j++){
                SXSSFCell cell_ij = row_i.createCell(j);
                cell_ij.setCellStyle(contentCellStyle);
                try {
//                    Field field = data.getClass().getDeclaredField(fieldList.get(j).getName());
                    Field field = fieldList.get(j);
                    field.setAccessible(true);
                    Object dataValue = field.get(data);

                    // 空值
                    if (dataValue == null){
                        continue;
                    }

                    // 时间类型需要格式化的
                    if (field.isAnnotationPresent(ExcelExportField.class) && !"".equals(field.getAnnotation(ExcelExportField.class).formatter())){
                        cell_ij.setCellValue(formatTimeObject(dataValue,field.getAnnotation(ExcelExportField.class).formatter()));
                        continue;
                    }

                    // 数值类型
                    if (dataValue instanceof Number) {
                        if(insertCellValueWithNumber(cell_ij,dataValue))
                            continue;
                    }

                    // 其他类型
                    cell_ij.setCellValue(dataValue.toString());
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return workbook;
    }

    /**
     * 生成工作簿并写入数据
     * @param workbook
     * @param dataList
     * @param c
     * @param sheetName
     * @param tableTitle
     * @param <T>
     * @return
     */
    public static <T> SXSSFWorkbook generateSheetAndWriteData(SXSSFWorkbook workbook, List<T> dataList, Class c, String sheetName, String tableTitle){
        List<Field> fieldList = getExcelExportField(c);
        // 创建需要的 sheet 和表头信息
        initSheetAndTable(workbook,fieldList,sheetName,tableTitle);
        // 写入数据
        return whiteDate(workbook,dataList,fieldList,sheetName);
    }

    /**
     * 获取类中要导出到 Excel 中的字段
     * 排除掉忽略的字段
     *
     * @param c
     * @return
     */
    public static List<Field> getExcelExportField(Class c){
        List<Field> fieldList = getAllFeild(c);
        List<Field> targetFeildList = fieldList.stream()
                .filter(field -> !field.isAnnotationPresent(ExcelExportField.class) || !field.getAnnotation(ExcelExportField.class).ignore())
                .sorted((f1, f2) -> {
                    ExcelExportField anno1 = f1.getAnnotation(ExcelExportField.class);
                    ExcelExportField anno2 = f2.getAnnotation(ExcelExportField.class);
                    if (anno1 == null && anno2 == null) return 0;
                    if (anno1 == null) return 1;
                    if (anno2 == null) return -1;
                    return Integer.compare(anno1.order(),anno2.order());
                })
                .collect(Collectors.toList());
        return targetFeildList;
    }

    /**
     * 获取类所有的字段，包括父类中的字段
     *
     * @param c
     * @return
     */
    public static List<Field> getAllFeild(Class c){
        if (c == null || c.equals(Object.class)){
            return new ArrayList<>();
        }
        List<Field> allFieldsInClass = Arrays.stream(c.getDeclaredFields()).collect(Collectors.toList());
        if (!c.getSuperclass().equals(Object.class)){
            allFieldsInClass.addAll(getAllFeild(c.getSuperclass()));
        }
        return allFieldsInClass;
    }

    /**
     * 格式化时间对象
     * @param timeObject 时间对象
     * @param pattern 格式化模式
     * @return
     */
    private static String formatTimeObject(Object timeObject, String pattern){
        String formattedTimeString = "";
        if (timeObject instanceof ZonedDateTime){
            formattedTimeString = ((ZonedDateTime)timeObject).format(DateTimeFormatter.ofPattern(pattern));
        }
        return formattedTimeString;
    }

    /**
     * 往 cell 中插入数值数据
     * @param cell 要插入数据的单元格
     * @param valueObject 要插入的数据对象
     * @return 是否插入成功
     */
    private static boolean insertCellValueWithNumber(SXSSFCell cell, Object valueObject){
        boolean insertSuccess = false;
        cell.setCellType(CellType.NUMERIC);
        if (valueObject instanceof Float){
            cell.setCellValue(((Number) valueObject).floatValue());
            insertSuccess = true;
        }else if(valueObject instanceof Long){
            cell.setCellValue(((Number) valueObject).longValue());
            insertSuccess = true;
        }else if (valueObject instanceof Integer) {
            cell.setCellValue(((Number) valueObject).intValue());
            insertSuccess = true;
        }else if (valueObject instanceof Double){
            cell.setCellValue(((Number) valueObject).doubleValue());
            insertSuccess = true;
        }else if (valueObject instanceof Short){
            cell.setCellValue(((Number) valueObject).shortValue());
            insertSuccess = true;
        }else if (valueObject instanceof Byte){
            cell.setCellValue(((Number) valueObject).byteValue());
            insertSuccess = true;
        }
        return insertSuccess;
    }
}
