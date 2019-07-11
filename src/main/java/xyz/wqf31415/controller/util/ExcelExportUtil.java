package xyz.wqf31415.controller.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import xyz.wqf31415.annotation.ExcelExportField;
import java.lang.reflect.Field;
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
     * 按给定的 字段列表 生成表格第一行
     *
     * @return
     */
    public static SXSSFWorkbook generateWorkbook(List<Field> targetFeildList, String sheetName){
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        SXSSFSheet sheet = workbook.createSheet(sheetName);
        CellStyle titleStyle = createTitleStyle(workbook);
        SXSSFRow row0 = sheet.createRow(0);
        for (int i = 0; i<targetFeildList.size();i++) {
            SXSSFCell cell_0_i = row0.createCell(i, CellType.STRING);
            cell_0_i.setCellStyle(titleStyle);
            Field f = targetFeildList.get(i);
            ExcelExportField anno = f.getAnnotation(ExcelExportField.class);
            String title = anno==null || anno.title() == null || "".equals(anno.title()) ? f.getName():anno.title();
            cell_0_i.setCellValue(title);
        }
        // 冻结第一行
        sheet.createFreezePane(0, 1, 0, 1);
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
     * 生成表格对象并写入数据
     *
     * @param dataList
     * @param c
     * @param <T>
     * @return
     */
    public static <T> SXSSFWorkbook whiteDate(List<T> dataList, Class c){
        List<Field> fieldList = getExcelExportField(c);
        SXSSFWorkbook book = generateWorkbook(fieldList,c.getSimpleName());
        SXSSFSheet sheet = book.getSheet(c.getSimpleName());
        CellStyle contentCellStyle = createContentCellStyle(book);
        for (int i = 0 ;i<dataList.size();i++){
            SXSSFRow row_i = sheet.createRow(i+1);
            T data = dataList.get(i);
            for (int j = 0; j<fieldList.size();j++){
                SXSSFCell cell_ij = row_i.createCell(j);
                cell_ij.setCellStyle(contentCellStyle);
                try {
//                    Field field = data.getClass().getDeclaredField(fieldList.get(j).getName());
                    Field field = fieldList.get(j);
                    field.setAccessible(true);
                    cell_ij.setCellValue(field.get(data).toString());
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return book;
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
        List<Field> result = Arrays.stream(c.getDeclaredFields()).collect(Collectors.toList());
        if (!c.getSuperclass().equals(Object.class)){
            result.addAll(getAllFeild(c.getSuperclass()));
        }
        return result;
    }
}
