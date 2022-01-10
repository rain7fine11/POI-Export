package com.wecool.poiexport;

import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.*;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * @author bowafterrain [mazhaoming@vip.qq.com]
 * @date 2021-12-29 20:28
 */
public class ExcelExport<E> implements Closeable {

    public static final int XLS_MAXIMUM_SIZE = 65536;
    private final HSSFWorkbook workbook = new HSSFWorkbook();

    private final HSSFSheet sheet = workbook.createSheet();

    private Class<E> dataClass;

    private Class<?> group = ExcelColumn.DefaultGroup.class;

    private Map<ExcelColumn, Field> columnMap;

    private boolean addSeqNo;

    private int rowNum;

    public ExcelExport() {
    }

    public ExcelExport(Class<E> dataClass) {
        this(dataClass, ExcelColumn.DefaultGroup.class);
    }

    public ExcelExport(Class<E> dataClass, Class<?> group) {
        this.group = group;
        this.initColumnMap(dataClass);
    }

    public Class<?> getGroup() {
        return group;
    }

    public void setGroup(Class<?> group) {
        if (rowNum > 0) {
            throw new IllegalStateException("cannot setGroup after addData");
        }
        this.group = group;
    }

    public boolean isAddSeqNo() {
        return addSeqNo;
    }

    public void setAddSeqNo(boolean addSeqNo) {
        this.addSeqNo = addSeqNo;
    }

    public void addDataList(List<E> dataList) throws IntrospectionException, InvocationTargetException, IllegalAccessException {
        if (dataList == null) {
            return;
        }
        int size = dataList.size();
        if (size == 0) {
            return;
        }
        if (dataClass == null) {
            this.initColumnMap(cast(dataList.get(0).getClass()));
        }
        if (rowNum == 0) {
            this.setHeader();
        }
        if (rowNum + size > XLS_MAXIMUM_SIZE) {
            throw new IndexOutOfBoundsException("xls file supports " + XLS_MAXIMUM_SIZE + " rows at most");
        }
        for (E data : dataList) {
            HSSFRow row = sheet.createRow(rowNum++);
            int column = 0;
            if (addSeqNo) {
                row.createCell(column++).setCellValue(rowNum - 1);
            }
            for (Map.Entry<ExcelColumn, Field> entry : columnMap.entrySet()) {
                PropertyDescriptor propertyDescriptor = new PropertyDescriptor(entry.getValue().getName(), dataClass);
                Object cellValue = propertyDescriptor.getReadMethod().invoke(data);
                row.createCell(column++).setCellValue(cellValue != null ? entry.getValue().getType() == Date.class
                        ? (DateFormatUtils.format((Date) cellValue, entry.getKey().datePattern())) : cellValue.toString()
                        : "");
            }
        }
    }

    public void addData(E data) throws IntrospectionException, InvocationTargetException, IllegalAccessException {
        this.addDataList(Collections.singletonList(data));
    }

    public void write(OutputStream outputStream) throws IOException {
        this.autoSetColumnWidth();
        workbook.write(outputStream);
    }

    @Override
    public void close() throws IOException {
        workbook.close();
    }

    private void initColumnMap(Class<E> dataClass) {
        this.dataClass = dataClass;
        columnMap = new TreeMap<>(Comparator.comparingInt(ExcelColumn::order));
        for (Field declaredField : dataClass.getDeclaredFields()) {
            for (ExcelColumn excelColumn : declaredField.getAnnotationsByType(ExcelColumn.class)) {
                if (excelColumn.group() == group && columnMap.put(excelColumn, declaredField) != null) {
                    throw new IllegalArgumentException(String.format("the field %s has multiple columns in the group %s",
                            declaredField, group));
                }
            }
        }
    }

    private void setHeader() {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        HSSFRow row = sheet.createRow(rowNum++);
        int column = 0;
        if (addSeqNo) {
            createCell(row, column++, cellStyle, "序号");
        }
        for (ExcelColumn excelColumn : columnMap.keySet()) {
            createCell(row, column++, cellStyle, excelColumn.name());
        }
    }

    private void createCell(HSSFRow row, int column, HSSFCellStyle cellStyle, String cellValue) {
        HSSFCell cell = row.createCell(column);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(cellValue);
    }

    private void autoSetColumnWidth() {
        for (int i = 0; i < columnMap.size(); i++) {
            sheet.autoSizeColumn(i);
            int columnWidth = sheet.getColumnWidth(i);
            if (columnWidth <= 43520) {
                sheet.setColumnWidth(i, (int) (Math.round(columnWidth * 1.5)));
            }
        }
    }

    @SuppressWarnings("unchecked")
    private static <T> T cast(Object object) {
        return (T) object;
    }
}
