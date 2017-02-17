package com.solverpeng.poi.utils.excel;

import com.solverpeng.poi.utils.Reflections;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.util.*;

/**
 * 导出Excel文件（导出“XLS”格式）
 * @author
 */
public class ExportExcel_bk {

    private static Logger log = LoggerFactory.getLogger(ExportExcel_bk.class);

    /**
     * 工作薄对象
     */
    private Workbook workbook;

    /**
     * 工作表对象
     */
    private Sheet sheet;

    /**
     * 样式列表
     */
    private Map<String, CellStyle> styles;

    /**
     * 当前行号
     */
    private int rowNum;

    /**
     * 注解列表（Object[]{ ExcelField, Field/Method }）
     */
    private List<Object[]> annotationList = new ArrayList<>();

    /**
     * @param titleName 表格标题，传“空值”，表示无标题
     * @param cls   实体对象，通过annotation.ExportField获取标题
     */
    public ExportExcel_bk(String titleName, Class<?> cls) {
        this(titleName, cls, 1);
    }

    /**
     * @param titleName 表格标题，传“空值”，表示无标题
     * @param clazz     实体对象，通过annotation.ExportField获取标题
     * @param type      导出类型（1:导出数据；2：导出模板）
     * @param groups    导入分组
     */
    public ExportExcel_bk(String titleName, Class<?> clazz, int type, int... groups) {
        // 将 @ExportField 标注的属性添加到 annotationList 中
        handleAnnotationFields(clazz, type, groups);
        // 将 @ExportField 标注的方法添加到 annotationList 中。
        // 注释原因：只留下一种方式，即将注解标注实体类的属性，减少了方法
        //handleAnnotationMethods(clazz, type, groups);
        // 将 @ExportField 标注的属性进行排序
        sortAnnotationFields();
        // 生成排序好的表头名称
        List<String> headerList = handleHeaderList(type);
        // 生成标题和表头信息
        initialize(titleName, headerList);
    }

    /**
     * @param titleName   表格标题，传“空值”，表示无标题
     * @param headers 表头数组
     */
    public ExportExcel_bk(String titleName, String[] headers) {
        initialize(titleName, Arrays.asList(headers));
    }

    /**
     * @param titleName      表格标题，传“空值”，表示无标题
     * @param headerList 表头列表
     */
    public ExportExcel_bk(String titleName, List<String> headerList) {
        initialize(titleName, headerList);
    }

    /**
     * 添加数据（通过annotation.ExportField添加数据）
     */
    public <E> ExportExcel_bk setDataList(List<E> list) {
        for (E e : list) {
            int column = 0;
            Row row = this.addRow();
            StringBuilder sb = new StringBuilder();
            for (Object[] objectArr : annotationList) {
                ExcelField excelField = (ExcelField) objectArr[0];
                Object val = null;
                // Get entity value
                try {
                    if (StringUtils.isNotBlank(excelField.value())) {
                        val = Reflections.invokeGetter(e, excelField.value());
                    } else {
                        if (objectArr[1] instanceof Field) {
                            val = Reflections.invokeGetter(e, ((Field) objectArr[1]).getName());
                        }
                        /*if (objectArr[1] instanceof Method) {
                            val = Reflections.invokeMethod(e, ((Method) objectArr[1]).getName(), new Class[]{}, new Object[]{});
                        }*/
                    }
                } catch (Exception ex) {
                    log.info(ex.toString());
                    val = "";
                }
                this.addCell(row, column++, val, excelField.align(), excelField.fieldType());
                sb.append(val).append(", ");
            }
            log.debug("Write success: [" + row.getRowNum() + "] " + sb.toString());
        }
        return this;
    }

    public Row addRow() {
        return sheet.createRow(rowNum++);
    }

    /**
     * @param row    添加的行
     * @param column 添加列号
     * @param val    添加值
     */
    public Cell addCell(Row row, int column, Object val) {
        return this.addCell(row, column, val, 0, Class.class);
    }

    /**
     * @param row    添加的行
     * @param column 添加列号
     * @param val    添加值
     * @param align  对齐方式（1：靠左；2：居中；3：靠右）
     */
    public Cell addCell(Row row, int column, Object val, int align, Class<?> fieldType) {
        Cell cell = row.createCell(column);
        CellStyle style = styles.get("data" + (align >= 1 && align <= 3 ? align : ""));
        CellStyle dateStyle = styles.get("date" + (align >= 1 && align <= 3 ? align : ""));
        Boolean isDate = false;
        try {
            if (val == null) {
                cell.setCellValue("");
            } else if (val instanceof String) {
                cell.setCellValue((String) val);
            } else if (val instanceof Integer) {
                cell.setCellValue((Integer) val);
            } else if (val instanceof Long) {
                cell.setCellValue((Long) val);
            } else if (val instanceof Double) {
                cell.setCellValue((Double) val);
            } else if (val instanceof Float) {
                cell.setCellValue((Float) val);
            } else if (val instanceof Date) {
                cell.setCellValue((Date) val);
                isDate = true;
            } else {
                if (fieldType != Class.class) {
                    cell.setCellValue((String) fieldType.getMethod("setValue", Object.class).invoke(null, val));
                } else {
                    cell.setCellValue((String) Class.forName(this.getClass().getName().replaceAll(this.getClass().getSimpleName(),
                            "fieldtype." + val.getClass().getSimpleName() + "Type")).getMethod("setValue", Object.class).invoke(null, val));
                }
            }
        } catch (Exception ex) {
            log.info("Set cell value [" + row.getRowNum() + "," + column + "] error: " + ex.toString());
            assert val != null;
            cell.setCellValue(val.toString());
        }
        if (isDate) {
            cell.setCellStyle(dateStyle);
        } else {
            cell.setCellStyle(style);
        }
        return cell;
    }

    /**
     * 输出到浏览器
     */
    public ExportExcel_bk writeFile(String fileName, HttpServletRequest request, HttpServletResponse response) throws IOException {
        fileName = fileName + ".xls";
        response.reset();
        response.setContentType("application/vnd.ms-excel");
        setFileDownloadHeader(request, response, fileName);

        OutputStream op = response.getOutputStream();
        workbook.write(op);
        op.close();
        return this;
    }

    /**
     * 输出到文件系统
     */
    public String write2File(String downLoadDir) throws IOException {
        String fileName = downLoadDir + "/" + System.currentTimeMillis() + ".xls";
        FileOutputStream fileOut = new FileOutputStream(fileName);
        workbook.write(fileOut);
        fileOut.close();
        return fileName;
    }

    private List<String> handleHeaderList(int type) {
        List<String> headerList = new ArrayList<>();
        for (Object[] os : annotationList) {
            String t = ((ExcelField) os[0]).title();
            // 如果是导出模板，则去掉表头中的注释，如 表头**表头注释
            if (type == 1 || type == 0) {
                String[] ss = StringUtils.split(t, "**", 2);
                if (ss.length == 2) {
                    t = ss[0];
                }
            }
            headerList.add(t);
        }
        return headerList;
    }

    private void sortAnnotationFields() {
        Collections.sort(annotationList, new Comparator<Object[]>() {
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExcelField) o1[0]).sort()).compareTo(((ExcelField) o2[0]).sort());
            }
        });
    }

    private void handleAnnotationFields(Class<?> clazz, int type, int[] groups) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null && (excelField.type() == 0 || excelField.type() == type)) {
                if (groups != null && groups.length > 0) {
                    boolean inGroup = false;
                    for (int group : groups) {
                        if (inGroup) {
                            break;
                        }
                        for (int excelFieldGroup : excelField.groups()) {
                            if (group == excelFieldGroup) {
                                inGroup = true;
                                annotationList.add(new Object[]{excelField, field});
                                break;
                            }
                        }
                    }
                } else {
                    annotationList.add(new Object[]{excelField, field});
                }
            }
        }
    }

    private void handleAnnotationMethods(Class<?> clazz, int type, int[] groups) {
        Method[] methods = clazz.getDeclaredMethods();
        for (Method method : methods) {
            ExcelField excelField = method.getAnnotation(ExcelField.class);
            if (excelField != null && (excelField.type() == 0 || excelField.type() == type)) {
                if (groups != null && groups.length > 0) {
                    boolean inGroup = false;
                    for (int g : groups) {
                        if (inGroup) {
                            break;
                        }
                        for (int efg : excelField.groups()) {
                            if (g == efg) {
                                inGroup = true;
                                annotationList.add(new Object[]{excelField, method});
                                break;
                            }
                        }
                    }
                } else {
                    annotationList.add(new Object[]{excelField, method});
                }
            }
        }
    }

    /**
     * @param titleName  表格标题，传“空值”，表示无标题
     * @param headerList 表头列表
     */
    private void initialize(String titleName, List<String> headerList) {
        this.workbook = new HSSFWorkbook();
        this.sheet = workbook.createSheet("sheet");
        this.styles = createStyles(workbook);

        // 创建表格标题
        buildExcelTitle(titleName, headerList.size());
        // 创建表头列表:表头可以为 表头**表头注释 的形式
        buildExcelHeader(headerList);

        log.debug("表格初始化成功!");
    }

    private void buildExcelTitle(String titleName, int headListSize) {
        if (StringUtils.isNotBlank(titleName)) {
            Row titleRow = sheet.createRow(rowNum++);
            titleRow.setHeightInPoints(30);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellStyle(styles.get("title"));
            titleCell.setCellValue(titleName);
            int rowNum = titleRow.getRowNum();
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowNum, rowNum, rowNum, headListSize - 1);
            sheet.addMergedRegion(cellRangeAddress);
        }
    }

    private void buildExcelHeader(List<String> headerList) {
        if (headerList == null) {
            throw new RuntimeException("表格表头不能为空!");
        }
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.setHeightInPoints(16);
        for (int i = 0; i < headerList.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(styles.get("header"));
            String[] headerAndComment = StringUtils.split(headerList.get(i), "**", 2);
            if (headerAndComment.length == 2) {
                cell.setCellValue(headerAndComment[0]);
                HSSFClientAnchor clientAnchor = new HSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6);
                Comment comment = this.sheet.createDrawingPatriarch().createCellComment(clientAnchor);
                comment.setString(new HSSFRichTextString(headerAndComment[1]));
                cell.setCellComment(comment);
            } else {
                cell.setCellValue(headerList.get(i));
            }
        }
        for (int i = 0; i < headerList.size(); i++) {
            int colWidth = sheet.getColumnWidth(i) * 2;
            sheet.setColumnWidth(i, colWidth < 3000 ? 3000 : colWidth);
        }
    }

    /**
     * 创建表格样式
     */
    private Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

        CellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style.setFont(titleFont);
        styles.put("title", style);

        style = wb.createCellStyle();
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_LEFT);
        styles.put("data1", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_CENTER);
        styles.put("data2", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        styles.put("data3", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        styles.put("header", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("yyyy-MM-dd"));
        styles.put("date", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("date"));
        style.setAlignment(CellStyle.ALIGN_LEFT);
        styles.put("date1", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("date"));
        style.setAlignment(CellStyle.ALIGN_CENTER);
        styles.put("date2", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("date"));
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        styles.put("date3", style);
        return styles;
    }

    /**
     * 根据当前用户的浏览器不同，对文件的名字进行不同的编码设置，从而解决不同浏览器下文件名中文乱码问题
     */
    private void setFileDownloadHeader(HttpServletRequest request, HttpServletResponse response, String fileName) {
        String encodedFileName;
        try {
            //中文文件名支持
            if (request.getHeader("User-Agent").toUpperCase().indexOf("MSIE") > 0) {//IE浏览器
                encodedFileName = URLEncoder.encode(fileName, "UTF-8");
            } else if (request.getHeader("User-Agent").toLowerCase().indexOf("firefox") > 0 || request.getHeader("User-Agent").toLowerCase().indexOf("opera") > 0) {//google,火狐浏览器
                encodedFileName = new String(fileName.getBytes(), "ISO8859-1");
            } else {
                encodedFileName = URLEncoder.encode(fileName, "UTF-8");//其他浏览器
            }

            response.setHeader("Content-Disposition", "attachment; filename=\"" + encodedFileName + "\"");//这里设置一下让浏览器弹出下载提示框，而不是直接在浏览器中打开
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
    }

}