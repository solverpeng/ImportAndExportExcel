package com.solverpeng.poi.utils.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 导出Excel文件（导出“XLS”格式）
 *
 * @author solverpeng
 */
public class ExportExcelUtil {

    private static Logger log = LoggerFactory.getLogger(ExportExcel.class);

    /**
     * Excel 导出
     *
     * @param titleName 标题名称
     * @param headers   标头
     * @param fillData  填充的数据
     * @param fileName  导出 Excel 文件名称
     * @author solverpeng
     */
    public static void downLoad4Excel(String titleName, List<String> headers, List<Object[]> fillData, String fileName,
                                      HttpServletRequest request, HttpServletResponse response) throws IOException {
        downLoad4Excel(titleName, headers, fillData, fileName, 0, request, response);
    }

    /**
     * @param align 数据对齐方式 1：左对齐 2：居中 3：右对齐
     * @author solverpeng
     */
    public static void downLoad4Excel(String titleName, List<String> headers, List<Object[]> fillData, String fileName,
                                      int align, HttpServletRequest request, HttpServletResponse response) throws IOException {
        downLoad4Excel(titleName, headers, fillData, fileName, align, "sheet", request, response);
    }

    /**
     * @param sheetName 表格名称
     * @author solverpeng
     */
    public static void downLoad4Excel(String titleName, List<String> headers, List<Object[]> fillData, String fileName,
                                      int align, String sheetName, HttpServletRequest request, HttpServletResponse response) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);

        // 0.获取样式
        Map<String, CellStyle> styles2 = createStyles2(workbook);

        // 1.设置标题
        buildTitleRow(titleName, sheet, styles2, headers.size());

        // 2.设置标头
        buildHeaderRow(headers, sheet, styles2);

        // 3.填充数据
        fillInToData(fillData, sheet, styles2, align, workbook);

        // 4.下载
        writeFile(workbook, fileName, request, response);
    }

    private static void buildTitleRow(String titleName, Sheet sheet, Map<String, CellStyle> styles2, int headerListSize) {
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(30);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellStyle(styles2.get("title"));
        titleCell.setCellValue(titleName);
        sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),
                titleRow.getRowNum(), titleRow.getRowNum(), headerListSize - 1));
    }

    private static void buildHeaderRow(List<String> headers, Sheet sheet, Map<String, CellStyle> styles2) {
        Row headerRow = sheet.createRow(1);
        headerRow.setHeightInPoints(16);
        for (int i = 0; i < headers.size(); i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellStyle(styles2.get("header"));
            headerCell.setCellValue(headers.get(i));
        }

        for (int i = 0; i < headers.size(); i++) {
            int colWidth = sheet.getColumnWidth(i) * 2;
            sheet.setColumnWidth(i, colWidth < 3000 ? 3000 : colWidth);
        }
    }

    private static void fillInToData(List<Object[]> list, Sheet sheet, Map<String, CellStyle> styles2, int align, Workbook workbook) {
        int rowNum = 2;// 填充数据从第3行开始，即 rowNum =2
        CellStyle dataStyle = styles2.get("data" + (align >= 1 && align <= 3 ? align : ""));

        for (Object[] objectArr : list) {
            Object[] objects = (Object[]) objectArr;
            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < objects.length; i++) {
                Cell cell = row.createCell(i);
                Object object = objects[i];
                if (object != null) {
                    if (object instanceof Date) {
                        DataFormat format = workbook.createDataFormat();
                        dataStyle.setDataFormat(format.getFormat("yyyy-MM-dd"));
                        cell.setCellValue((Date) object);
                    } else {
                        cell.setCellValue(object.toString());
                    }
                }
                cell.setCellStyle(dataStyle);
            }
        }
    }

    private static void writeFile(Workbook wb, String fileName, HttpServletRequest request, HttpServletResponse response) throws IOException {
        //导出excel文档
        OutputStream op = null;
        fileName = fileName + ".xls";
        response.reset();
        response.setContentType("application/vnd.ms-excel");
        setFileDownloadHeader(request, response, fileName);

        op = response.getOutputStream();
        wb.write(op);
        op.close();
    }

    //根据当前用户的浏览器不同，对文件的名字进行不同的编码设置，从而解决不同浏览器下文件名中文乱码问题
    private static void setFileDownloadHeader(HttpServletRequest request, HttpServletResponse response, String fileName) {
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

    private static Map<String, CellStyle> createStyles2(Workbook wb) {
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

        return styles;
    }


}