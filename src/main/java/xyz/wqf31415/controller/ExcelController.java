package xyz.wqf31415.controller;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import xyz.wqf31415.entity.Student;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

/**
 * Created by Administrator on 2018/2/26.
 *
 * @author WeiQuanfu
 */
@RestController
@RequestMapping("/excel")
public class ExcelController {

    /**
     * 下载excel表格
     *
     * @param response
     */
    @RequestMapping("/create-xls")
    public void download(HttpServletResponse response){
        HSSFWorkbook workbook = new HSSFWorkbook();
        CreationHelper creationHelper = workbook.getCreationHelper();
        // Excel 表格的工作表名称长度不能超过31个字符，不能包含 0x000 0x003 : \ * / ? : [ ]
        // 可以用 POI 提供的工具类创建正确的工作表名，非法字符将被替换成空格，以下代码将返回 s h e e t 1
        String sheetName = WorkbookUtil.createSafeSheetName("[s/h*e:e?t\\1]");
        System.out.println("sheetName:"+sheetName);
        HSSFSheet sheet = workbook.createSheet(sheetName);
        HSSFRow row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("姓名");
        row0.createCell(1).setCellValue("年龄");
        row0.createCell(2).setCellValue("性别");
        row0.createCell(3).setCellValue("电话");

        RichTextString richTextString = creationHelper.createRichTextString("Date日期");
        row0.createCell(4).setCellValue(richTextString);

        row0.createCell(5).setCellValue("Calendar日期");
        row0.createCell(6).setCellValue("单元格类型");

        HSSFRow row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("张三");
        row1.createCell(1).setCellValue(18);
        row1.createCell(2).setCellValue("男");
        row1.createCell(3).setCellValue("87878787");

        // 给单元格的时间数据设置显示格式
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
        HSSFCell cell_1_4 = row1.createCell(4);
        cell_1_4.setCellValue(new Date());
        cell_1_4.setCellStyle(cellStyle);

        HSSFCell cell_1_5 = row1.createCell(5);
        cell_1_5.setCellValue(Calendar.getInstance());
        cell_1_5.setCellStyle(cellStyle);

        row1.createCell(6).setCellType(CellType.ERROR);

        String excelFileName = "details.xls";
        try {
            ServletOutputStream out = response.getOutputStream();
            response.reset();
            response.setHeader("Content-disposition", "attachment; filename=" + excelFileName);
            response.setContentType("application/msexcel");
            workbook.write(out);
            out.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 将实体List导出为 Excel
     *
     * @param response
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     * @throws IOException
     */
    @RequestMapping("/students")
    public void downStudentExcel(HttpServletResponse response) throws NoSuchFieldException, IllegalAccessException, IOException {
        // 要导出的数据
        Student s1 = new Student("张三",18,"男");
        Student s2 = new Student("李四",19,"女");
        List<Student> studentList = new ArrayList<>();
        studentList.add(s1);
        studentList.add(s2);

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("学生名单");
        int rowNumber = 0;
        HSSFRow row0 = sheet.createRow(rowNumber++);

        Class studentClass = Student.class;
        Field[] fields = studentClass.getDeclaredFields();
        // 写入首行标题
        for (int i = 0; i<fields.length;i++){
            System.out.println("名称:"+fields[i].getName());
            row0.createCell(i).setCellValue(fields[i].getName());
        }
        // 写入内容
        for (Student stu : studentList){
            HSSFRow row = sheet.createRow(rowNumber++);
            for (int i = 0; i<fields.length;i++){
                Field f = stu.getClass().getDeclaredField(fields[i].getName());
                f.setAccessible(true);
                row.createCell(i).setCellValue(f.get(stu).toString());
            }
        }

        // 导出 excel
        String excelFileName = "student.xls";
        ServletOutputStream out = response.getOutputStream();
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=" + excelFileName);
        response.setContentType("application/msexcel");
        workbook.write(out);
        out.close();
    }

    /**
     * 上传 Excel 文件，程序读取内容后以字符串形式返回
     *
     * @param file
     * @return
     * @throws IOException
     */
    @RequestMapping("/upload-parse")
    public String uploadExcelFile(MultipartFile file) throws IOException {
        String result = "";
        InputStream stream = file.getInputStream();
        HSSFWorkbook workbook = new HSSFWorkbook(stream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        // 获取最后一行的索引
        int rowNum = sheet.getLastRowNum();
        for (int i = 0; i <= rowNum; i++){
            HSSFRow row = sheet.getRow(i);
            // 获取单元格数量，即最后一个单元格索引+1
            int cellNum = row.getLastCellNum();
            for (int j = 0; j < cellNum; j++){
                HSSFCell cell = row.getCell(j);
                if (cell != null) {
                    result += cell.getStringCellValue() + "\t";
                }
            }
            result += "\n";
        }
        System.out.println(result);
        return result;
    }

    /**
     * 导入学生名单，解析数据后以 json 格式返回
     *
     * @param file
     * @return
     * @throws IOException
     */
    @RequestMapping("/upload-students")
    public String uploadStudents(MultipartFile file) throws IOException {
        InputStream stream = file.getInputStream();
        HSSFWorkbook workbook = new HSSFWorkbook(stream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        int rowNum = sheet.getLastRowNum();
        JSONArray array = new JSONArray();
        for (int i = 1; i <= rowNum; i++){
            HSSFRow row = sheet.getRow(i);
            Student student = new Student();
            // 第一行是表头，从第二行开始是学生信息
            student.setName(row.getCell(0).getStringCellValue());
            student.setAge(Integer.parseInt(row.getCell(1).getStringCellValue()));
            student.setGender(row.getCell(2).getStringCellValue());
            JSONObject jsonObject = new JSONObject(student);
            array.put(jsonObject);
        }
        return array.toString();
    }

    /**
     * 导出 xlsx 文件
     * @param response
     * @throws IOException
     */
    @RequestMapping("/create-xlsx")
    public void createXlsx(HttpServletResponse response) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        addContent(workbook);
        String excelFileName = "test.xlsx";
        OutputStream stream = response.getOutputStream();
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=" + excelFileName);
        response.setContentType("application/msexcel");
        workbook.write(stream);
        stream.close();
    }

    @RequestMapping("/open-xlsx")
    public String openXlsx(MultipartFile file) throws IOException {
        InputStream inputStream = file.getInputStream();
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        return parseWorkbook(workbook);
    }

    /**
     * 解析 Excel 工作簿内容，以json格式返回
     *
     * @param workbook
     * @return
     */
    private String parseWorkbook(Workbook workbook) {
        int sheetNum = workbook.getNumberOfSheets();
        JSONArray sheetArr = new JSONArray();
        for (int i = 0; i < sheetNum; i++){
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = workbook.getSheetName(i);
            JSONObject sheetJson = new JSONObject();
            sheetJson.put("sheetName",sheetName);
            JSONArray rowArr = new JSONArray();
            int rowNum = sheet.getLastRowNum() + 1;
            for (int j = 0; j < rowNum; j++){
                Row row = sheet.getRow(j);
                JSONObject rowJson = new JSONObject();
                int cellNum = row.getLastCellNum();
                for (int k = 0; k < cellNum; k++){
                    Cell cell = row.getCell(k);
                    if(cell != null) {
                        CellType cellType = cell.getCellTypeEnum();
                        switch (cellType) {
                            case NUMERIC:
                                // 数字类型数据可能是数字，也可能是日期
                                if (DateUtil.isCellDateFormatted(cell)){
                                    rowJson.put("" + (char) (j + 65) + (k + 1), cell.getDateCellValue());
                                }else {
                                    rowJson.put("" + (char) (j + 65) + (k + 1), cell.getNumericCellValue());
                                }
                                break;
                            case STRING:
                                rowJson.put("" + (char) (j + 65) + (k + 1), cell.getStringCellValue());
                                break;
                            case BOOLEAN:
                                rowJson.put("" + (char) (j + 65) + (k + 1), cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                rowJson.put("" + (char) (j + 65) + (k + 1), cell.getCellFormula());
                                break;
                            case BLANK:
                                rowJson.put("" + (char) (j + 65) + (k + 1), "");
                                break;
                            default:
                                rowJson.put("" + (char) (j + 65) + (k + 1), "");
                        }
                    }else {
                        rowJson.put("" + (char) (j + 65) + (k + 1), "");
                    }
                }
                rowArr.put(rowJson);
            }
            sheetJson.put("data",rowArr);
            sheetArr.put(sheetJson);
        }
        return sheetArr.toString();
    }

    /**
     * 给 workbook 对象添加内容数据
     * @param workbook
     */
    private void addContent(Workbook workbook){
        CreationHelper helper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet();
        Row row_0 = sheet.createRow(0);
        row_0.createCell(0).setCellValue("Double");
        row_0.createCell(1).setCellValue("Date");
        row_0.createCell(2).setCellValue("Calendar");
        row_0.createCell(3).setCellValue("String");
        row_0.createCell(4).setCellValue("Boolean");
        row_0.createCell(5).setCellValue("RichTextString");

        Row row_1 = sheet.createRow(1);
        Cell cell_1_0 = row_1.createCell(0);
        Double d = 3.14;
        cell_1_0.setCellValue(d);

        Cell cell_1_1 = row_1.createCell(1);
        Date date = new Date();
        cell_1_1.setCellValue(date);

        Cell cell_1_2 = row_1.createCell(2);
        Calendar calendar = Calendar.getInstance();
        cell_1_2.setCellValue(calendar);

        Cell cell_1_3 = row_1.createCell(3);
        String str = "String";
        cell_1_3.setCellValue(str);

        Cell cell_1_4 = row_1.createCell(4);
        boolean b = true;
        cell_1_4.setCellValue(b);

        Cell cell_1_5 = row_1.createCell(5);
        RichTextString richTextString = helper.createRichTextString("富文本");
        cell_1_5.setCellValue(richTextString);
    }

    /**
     * 创建含有对齐方式的xlsx文件
     * @param response
     * @throws IOException
     */
    @RequestMapping("/create-align")
    public void createAlignedExcel(HttpServletResponse response) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("alignSheet");

        Row row_0 = sheet.createRow(0);
        Cell cell_0_0 = row_0.createCell(0);
        CellStyle cellStyle0 = workbook.createCellStyle();
        cellStyle0.setAlignment(HorizontalAlignment.LEFT);
        cellStyle0.setVerticalAlignment(VerticalAlignment.TOP);
        cell_0_0.setCellStyle(cellStyle0);
        cell_0_0.setCellValue("左上");

        Cell cell_0_1 = row_0.createCell(1);
        CellStyle cellStyle1 = workbook.createCellStyle();
        cellStyle1.setAlignment(HorizontalAlignment.CENTER);
        cellStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
        cell_0_1.setCellStyle(cellStyle1);
        cell_0_1.setCellValue("中间");

        Cell cell_0_2 = row_0.createCell(2);
        CellStyle cellStyle2 = workbook.createCellStyle();
        cellStyle2.setAlignment(HorizontalAlignment.RIGHT);
        cellStyle2.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cell_0_2.setCellStyle(cellStyle2);
        cell_0_2.setCellValue("右下");

        String excelFileName = "align.xlsx";
        ServletOutputStream stream = response.getOutputStream();
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=" + excelFileName);
        response.setContentType("application/msexcel");
        workbook.write(stream);
        stream.close();
    }

    /**
     * 创建含有边框样式的xls文件
     * @param response
     */
    @RequestMapping("/create-border")
    public void createXlsWithBorder(HttpServletResponse response) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("border");
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);
        cell.setCellValue("边框");
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.DOUBLE);
        cellStyle.setTopBorderColor(IndexedColors.BLUE.getIndex());
        cellStyle.setBorderRight(BorderStyle.DASH_DOT);
        cellStyle.setRightBorderColor(IndexedColors.RED.getIndex());
        cellStyle.setBorderBottom(BorderStyle.DOTTED);
        cellStyle.setBottomBorderColor(IndexedColors.GREEN.getIndex());
        cellStyle.setBorderLeft(BorderStyle.DASH_DOT_DOT);
        cellStyle.setLeftBorderColor(IndexedColors.YELLOW.getIndex());
        cell.setCellStyle(cellStyle);

        String excelFileName = "border.xls";
        ServletOutputStream stream = response.getOutputStream();
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=" + excelFileName);
        response.setContentType("application/msexcel");
        workbook.write(stream);
        stream.close();
    }
}
