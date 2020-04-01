package com.tony.demo.excel_helper;

import com.tony.demo.excel_helper.annotation.ExcelProperty;
import com.tony.demo.excel_helper.consts.ExcelType;
import com.tony.demo.excel_helper.exception.ExcelParseException;
import com.tony.demo.excel_helper.utils.StringUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * 功能说明： EXCEL操作助手
 * Created by TonyZeng on 2017/9/4.
 */
public class ExcelWriter {
    public ExcelWriter() {
    }

    /**
     * @param excelType
     * @param sheetName character count must be greater than or equal to 1 and less than or equal to 31.
     * @param data
     * @param classType
     * @param <T>
     * @return
     */
    public <T> String write(ExcelType excelType, String sheetName, List<T> data, Class<T> classType, OutputStream outputStream) {
        Workbook workbook = null;
        try {
            //Get　excel headers.
            List<ExcelEntityField> headers = _getExcelEntityFieldList(classType);


            workbook = excelType == ExcelType.xlsx ? new XSSFWorkbook() : new HSSFWorkbook();
            Sheet sheet = workbook.createSheet(sheetName);
            //产生表格标题行
            Row heardRow = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = heardRow.createCell(i);
                cell.setCellValue(headers.get(i).getColumnName());
            }
            //遍历集合数据，产生数据行
            for (int i = 0; i < data.size(); i++) {
                //数据行从第二行开始，即index=1
                Row dataRow = sheet.createRow(i + 1);
                List<String> values = _getExcelEntityValueList(headers, data.get(i));
                for (int j = 0; j < values.size(); j++) {
                    dataRow.createCell(j).setCellValue(values.get(j));
                }
            }

        } catch (SecurityException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            //清理资源
        }
        try {
            if (workbook != null) {
                workbook.write(outputStream);
            }
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "";
    }


    /**
     * 获取Excel实体类中的填充字段
     */
    private <T> List<ExcelEntityField> _getExcelEntityFieldList(Class<T> classType) throws ExcelParseException {
        List<ExcelEntityField> eefs = new ArrayList<>();
        // 遍历所有字段
        Field[] allFields = classType.getDeclaredFields();
        for (Field field : allFields) {
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            // 只对含有@ExcelProperty注解的字段进行赋值
            if (excelProperty != null) {
                ExcelEntityField excelEntityField = new ExcelEntityField();
                excelEntityField.setField(field);
                excelEntityField.setColumnName(excelProperty.value().trim());
                excelEntityField.setRequired(excelProperty.required());
                excelEntityField.setAnnotation(excelProperty);
                eefs.add(excelEntityField);
            }
        }
        return eefs;
    }

    /**
     * 获取Excel实体类中的填充字段
     */
    private <T> List<String> _getExcelEntityValueList(List<ExcelEntityField> entityFields, T data) throws ExcelParseException {
        List<String> result = new ArrayList<>();
        try {
            for (ExcelEntityField excelEntityField : entityFields) {
                Object value = data.getClass().getDeclaredMethod("get#FIELD".replace("#FIELD", StringUtil.toCapitalizeCamelCase(excelEntityField.getField().getName()))).invoke(data);
                result.add(value != null ? value.toString() : "");
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return result;
    }
}
