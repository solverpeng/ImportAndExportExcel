package com.solverpeng.poi.utils.excel;

import com.solverpeng.poi.utils.Reflections;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 导入Excel文件（支持“XLS” 和 "XLSX"格式的导入）
 *
 * @author solverpeng
 */
public class ImportExcel {

    private static Logger log = LoggerFactory.getLogger(ImportExcel.class);

    /**
     * 工作薄对象
     */
    private Workbook workbook;

    /**
     * 工作表对象
     */
    private Sheet sheet;

    /**
     * 标题行号
     */
    private int headerNum;

    /**
     * 注解列表（Object[]{ ExcelField, Field/Method }）
     */
    private List<Object[]> annotationList;

    /**
     * @param headerNum 标题行号，数据行号=标题行号+1
     */
    public ImportExcel(String fileName, int headerNum) throws IOException {
        this(new File(fileName), headerNum);
    }

    /**
     * @param headerNum 标题行号，数据行号=标题行号+1
     */
    public ImportExcel(File file, int headerNum) throws IOException {
        this(file, headerNum, 0);
    }

    /**
     * @param headerNum  标题行号，数据行号=标题行号+1
     * @param sheetIndex 工作表编号
     */
    public ImportExcel(String fileName, int headerNum, int sheetIndex) throws IOException {
        this(new File(fileName), headerNum, sheetIndex);
    }

    /**
     * @param headerNum  标题行号，数据行号=标题行号+1
     * @param sheetIndex 工作表编号
     */
    public ImportExcel(File file, int headerNum, int sheetIndex) throws IOException {
        this(file.getName(), new FileInputStream(file), headerNum, sheetIndex);
    }

    /**
     * 构造函数
     *
     * @param headerNum  标题行号，数据行号=标题行号+1
     * @param sheetIndex 工作表编号
     */
    public ImportExcel(MultipartFile multipartFile, int headerNum, int sheetIndex) throws IOException {
        this(multipartFile.getOriginalFilename(), multipartFile.getInputStream(), headerNum, sheetIndex);
    }

    /**
     * @param headerNum  标题行号，数据行号=标题行号+1
     * @param sheetIndex 工作表编号
     */
    public ImportExcel(String fileName, InputStream is, int headerNum, int sheetIndex) throws IOException {
        annotationList = new ArrayList<>();
        if (StringUtils.isBlank(fileName)) {
            throw new RuntimeException("导入文档为空!");
        } else if (fileName.toLowerCase().endsWith("xls")) {
            this.workbook = new HSSFWorkbook(is);
        } else if (fileName.toLowerCase().endsWith("xlsx")) {
            this.workbook = new XSSFWorkbook(is);
        } else {
            throw new RuntimeException("文档格式不正确!");
        }
        if (this.workbook.getNumberOfSheets() < sheetIndex) {
            throw new RuntimeException("文档中没有工作表!");
        }
        this.sheet = this.workbook.getSheetAt(sheetIndex);
        this.headerNum = headerNum;
        log.debug("Initialize success.");
    }

    /**
     * 获取导入数据列表
     *
     * @param clazz  导入对象类型
     * @param groups 导入分组
     */
    public <E> List<E> getDataList(Class<E> clazz, int... groups) throws IllegalAccessException, InstantiationException {
        // 将 @ExportField 标注的属性添加到 annotationList 中
        handleAnnotationFields(clazz, groups);
        // 将 @ExportField 标注的属性进行排序
        sortAnnotationFields();

        return handleDataList(clazz, groups);
    }

    /**
     * 获取行对象
     */
    public Row getRow(int rowNum) {
        return this.sheet.getRow(rowNum);
    }

    /**
     * 获取数据行号
     */
    public int getDataRowNum() {
        return headerNum + 1;
    }

    /**
     * 获取最后一个数据行号
     */
    public int getLastDataRowNum() {
        return this.sheet.getLastRowNum() + headerNum;
    }

    /**
     * 获取最后一个列号
     */
    public int getLastCellNum() {
        return this.getRow(headerNum).getLastCellNum();
    }

    /**
     * @param row    获取的行
     * @param column 获取单元格列号
     * @return 单元格值
     */
    public Object getCellValue(Row row, int column) {
        Object val = "";
        try {
            Cell cell = row.getCell(column);
            if (cell != null) {
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    val = cell.getNumericCellValue();
                } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    val = cell.getStringCellValue();
                } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    val = cell.getCellFormula();
                } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
                    val = cell.getBooleanCellValue();
                } else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
                    val = cell.getErrorCellValue();
                }
            }
        } catch (Exception e) {
            return val;
        }
        return val;
    }

    private <E> List<E> handleDataList(Class<E> clazz, int[] groups) throws IllegalAccessException, InstantiationException {
        List<E> dataList = new ArrayList<>();
        for (int i = this.getDataRowNum(); i < this.getLastDataRowNum(); i++) {
            E e = (E) clazz.newInstance();
            int column = 0;
            Row row = this.getRow(i);
            StringBuilder sb = new StringBuilder();
            for (Object[] objectArr : annotationList) {
                Object cellValue = this.getCellValue(row, column++);
                if (cellValue != null) {
                    ExcelField excelField = (ExcelField) objectArr[0];
                    // Get param type and type cast
                    Class<?> valType = Class.class;
                    if (objectArr[1] instanceof Field) {
                        valType = ((Field) objectArr[1]).getType();
                    }
                    try {
                        if (valType == String.class) {
                            cellValue = String.valueOf(cellValue.toString());
                        } else if (valType == Integer.class) {
                            cellValue = Double.valueOf(cellValue.toString()).intValue();
                        } else if (valType == Long.class) {
                            cellValue = Double.valueOf(cellValue.toString()).longValue();
                        } else if (valType == Double.class) {
                            cellValue = Double.valueOf(cellValue.toString());
                        } else if (valType == Float.class) {
                            cellValue = Float.valueOf(cellValue.toString());
                        } else if (valType == Date.class) {
                            cellValue = DateUtil.getJavaDate((Double) cellValue);
                        } else {
                            if (excelField.fieldType() != Class.class) {
                                cellValue = excelField.fieldType().getMethod("getValue", String.class)
                                        .invoke(null, cellValue.toString());
                            } else {
                                String replacement = "fieldtype." + valType.getSimpleName() + "Type";
                                cellValue = Class.forName(this.getClass().getName()
                                        .replaceAll(this.getClass().getSimpleName(), replacement))
                                        .getMethod("getValue", String.class)
                                        .invoke(null, cellValue.toString());
                            }
                        }
                    } catch (Exception ex) {
                        log.info("Get cell value [" + i + "," + column + "] error: " + ex.toString());
                        cellValue = null;
                    }
                    if (objectArr[1] instanceof Field) {
                        Reflections.invokeSetter(e, ((Field) objectArr[1]).getName(), cellValue);
                    }
                }
                sb.append(cellValue).append(", ");
            }
            dataList.add(e);
            log.debug("Read success: [" + i + "] " + sb.toString());
        }
        return dataList;
    }

    private void handleAnnotationFields(Class<?> clazz, int[] groups) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null && (excelField.type() == 3 || excelField.type() == 0)) {
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

    private void sortAnnotationFields() {
        Collections.sort(annotationList, new Comparator<Object[]>() {
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExcelField) o1[0]).sort()).compareTo(((ExcelField) o2[0]).sort());
            }
        });
    }

}
