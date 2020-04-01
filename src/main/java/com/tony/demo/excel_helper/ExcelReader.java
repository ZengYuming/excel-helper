package com.tony.demo.excel_helper;

import com.tony.demo.excel_helper.annotation.ExcelEntity;
import com.tony.demo.excel_helper.annotation.ExcelProperty;
import com.tony.demo.excel_helper.exception.ExcelContentInvalidException;
import com.tony.demo.excel_helper.exception.ExcelParseException;
import com.tony.demo.excel_helper.exception.ExcelRegexpValidFailedException;
import com.tony.demo.excel_helper.extension.ExcelRule;
import com.tony.demo.excel_helper.extension.ExcelFieldType;
import com.tony.demo.excel_helper.utils.StringUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 功能说明： EXCEL操作助手
 * 操作说明：
 * 使用方法很简单，只需要使用静态方法readExcel即可。
 * <code>
 * ExcelHelper eh = ExcelHelper.readExcel("excel文件名");
 * </code>
 * 如果要读取Excel中的标题栏有哪些
 * <code>eh.getHeaders()</code>
 * 如果要读取Excel中的数区域
 * <code>eh.getDatas()</code>
 * 读取到的数据按照Excel中存放的行列形式存放在二维数组中
 * 如果需要转换为实体列表的话
 * <code>eh.toEntitys(实体.class)</code>
 * 注意到是，实体类必须含有@ExcelEntity注解，同时需要用到的属性字段上需要
 * 用@ExcelProperty标注。
 * <p>
 * Created by TonyZeng on 2017/9/4.
 */
public class ExcelReader {
    // 最小列数目
    final public static int MIN_ROW_COLUMN_COUNT = 1;
    // 列索引
    private int lastColumnIndex;
    // 从Excel中读取的标题栏
    private String[] headers = null;

    public String[] getHeaders() {
        return headers;
    }

    // 从Excel中读取的数据
    private String[][] datas = null;

    public String[][] getDatas() {
        return datas;
    }

    // 规则对象缓存
    @SuppressWarnings("rawtypes")
    private static Map<String, ExcelRule> rulesCache = new HashMap<String, ExcelRule>();
    // 自定义字段类型对象缓存
    @SuppressWarnings("rawtypes")
    private static List<Class<? extends ExcelFieldType>> userDefinedType = new ArrayList<Class<? extends ExcelFieldType>>();

    private ExcelReader() {
    }

    /**
     * 注册新字段类型
     *
     * @param type
     * @throws ExcelParseException
     */
    public static void registerNewType(@SuppressWarnings("rawtypes") Class<? extends ExcelFieldType> type)
            throws ExcelParseException {
        if (!userDefinedType.contains(type)) {
            userDefinedType.add(type);
        }
    }

    /**
     * 读取Excel内容
     *
     * @param excelFilename
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    public static ExcelReader readExcel(String excelFilename) throws InvalidFormatException, IOException {
        return readExcel(excelFilename, 0);
    }

    /**
     * 读取Excel内容
     *
     * @param excelFilename
     * @param sheetIndex
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    public static ExcelReader readExcel(String excelFilename, int sheetIndex) throws InvalidFormatException, IOException {
        return readExcel(new File(excelFilename), sheetIndex);
    }

    /**
     * 读取Excel内容
     *
     * @param file
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    public static ExcelReader readExcel(File file) throws InvalidFormatException, IOException {
        return readExcel(file, 0);
    }

    /**
     * 读取Excel内容
     *
     * @param file
     * @param sheetIndex
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    public static ExcelReader readExcel(File file, int sheetIndex) throws InvalidFormatException, IOException {
        // 读取Excel工作薄
        Workbook wb = WorkbookFactory.create(file);
        return _readExcel(wb, sheetIndex);
    }

    /**
     * 从文件流读取Excel
     *
     * @param ins
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    public static ExcelReader readExcel(InputStream ins) throws InvalidFormatException, IOException {
        return readExcel(ins, 0);
    }

    /**
     * 读取Excel内容（从文件流）
     *
     * @param ins
     * @param sheetIndex
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    public static ExcelReader readExcel(InputStream ins, int sheetIndex) throws InvalidFormatException, IOException {
        return _readExcel(WorkbookFactory.create(ins), sheetIndex);
    }

    /**
     * 转换为实体
     */
    public <T> List<T> toEntitys(Class<T> classType) throws ExcelParseException, ExcelContentInvalidException, ExcelRegexpValidFailedException {
        // 如果实体没有@ExcelEntity，则不允许继续操作
        ExcelEntity excelEntity = classType.getAnnotation(ExcelEntity.class);
        if (excelEntity == null) {
            throw new ExcelParseException("Excel实体必须打上@ExcelEntity annotation!");
        }
        // 创建实体对象集
        List<T> entitys = new ArrayList<>();
        try {
            // 遍历提交的数据行，依次填充到创建的实体对象中
            for (String[] data : datas) {
                T obj = classType.newInstance();
                // 遍历实体对象的实体字段，通过反射为实体字段赋值
                for (ExcelEntityField excelEntityField : _getExcelEntityFieldList(classType)) {
                    // 实体数据填充
                    try {
                        //调用entity的set方法，进行设置值
                        obj.getClass()
                                .getDeclaredMethod("set#FIELD".replace("#FIELD", StringUtil.toCapitalizeCamelCase(excelEntityField.getField().getName())), excelEntityField.getField().getType())
                                .invoke(obj, _getFieldValue(data[excelEntityField.getIndex()], excelEntityField));
                    } catch (ExcelParseException e) {
                        if (excelEntityField.isRequired()) {
                            throw new ExcelParseException("字段" + excelEntityField.getColumnName() + "出错!", e);
                        }
                    } catch (ExcelContentInvalidException e) {
                        if (excelEntityField.isRequired()) {
                            throw e;
                        }
                    } catch (NullPointerException e) {
                        if (excelEntityField.isRequired()) {
                            throw new ExcelParseException("字段" + excelEntityField.getColumnName() + "出错!", e);
                        }
                    }
                }
                entitys.add(obj);
            }
        } catch (InstantiationException e1) {
            throw new ExcelParseException(e1);
        } catch (IllegalAccessException e1) {
            throw new ExcelParseException(e1);
        } catch (NoSuchMethodException e1) {
            throw new ExcelParseException(e1);
        } catch (SecurityException e1) {
            throw new ExcelParseException(e1);
        } catch (IllegalArgumentException e) {
            throw new ExcelParseException(e);
        } catch (InvocationTargetException e) {
            throw new ExcelParseException(e);
        } catch (Exception e) {
            throw new ExcelParseException(e);
        }

        return entitys;
    }


    /**
     * 读取Excel内容
     *
     * @param wb
     * @param sheetIndex
     * @return
     */
    private static ExcelReader _readExcel(Workbook wb, int sheetIndex) {
        ExcelReader eh = new ExcelReader();
        // 遍历Excel Sheet， 依次读取里面的内容
        if (sheetIndex > wb.getNumberOfSheets()) {
            return null;
        }
        Sheet sheet = wb.getSheetAt(sheetIndex);
        // 遍历表格的每一行
        int rowStart = sheet.getFirstRowNum();
        // 最小行数为1行
        int rowEnd = sheet.getLastRowNum();
        // 读取EXCEL标题栏
        eh._parseExcelHeader(sheet.getRow(0));
        // 读取EXCEL数据区域内容
        eh._parseExcelData(sheet, rowStart + 1, rowEnd);
        return eh;
    }

    /**
     * 解析EXCEL标题栏
     *
     * @param row
     */
    private void _parseExcelHeader(Row row) {
        lastColumnIndex = Math.max(row.getLastCellNum(), MIN_ROW_COLUMN_COUNT);
        headers = new String[lastColumnIndex];
        // 初始化headers，每一列的标题
        for (int columnIndex = 0; columnIndex < lastColumnIndex; columnIndex++) {
            Cell cell = row.getCell(columnIndex, Row.RETURN_BLANK_AS_NULL);
            headers[columnIndex] = _getCellValue(cell).trim();
        }
    }

    /**
     * 列名在列标题中的索引
     *
     * @param columnName
     * @return
     */
    private int _indexOfHeader(String columnName) {
        for (int i = 0; i < headers.length; i++) {
            if (headers[i].equals(columnName)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 解析EXCEL数据区域内容
     *
     * @param sheet
     * @param rowStart
     * @param rowEnd
     */
    private void _parseExcelData(Sheet sheet, int rowStart, int rowEnd) {
        datas = new String[rowEnd][lastColumnIndex];
        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            int rowNumber = rowIndex - rowStart;
            // 读取遍历每一行中的每一列
            for (int columnIndex = 0; columnIndex < lastColumnIndex; columnIndex++) {
                Cell cell = row.getCell(columnIndex, Row.RETURN_BLANK_AS_NULL);
                String value = _getCellValue(cell).trim();
                datas[rowNumber][columnIndex] = value;
            }
        }
    }

    /**
     * 读取每个单元格中的内容
     *
     * @param cell
     * @return
     */
    private String _getCellValue(Cell cell) {
        // 如果单元格为空的，则返回空字符串
        if (cell == null) {
            return "";
        }

        // 根据单元格类型，以不同的方式读取单元格的值
        String value = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").format(cell.getDateCellValue());
                } else {
                    value = (long) cell.getNumericCellValue() + "";
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                break;
            case Cell.CELL_TYPE_FORMULA:
                value = cell.getCellFormula();
                break;
            default:
        }
        return value;
    }

    /**
     * 获取Excel实体类中的填充字段
     */
    public <T> List<ExcelEntityField> _getExcelEntityFieldList(Class<T> classType) throws ExcelParseException {
        List<ExcelEntityField> eefs = new ArrayList<>();
        // 遍历所有字段
        Field[] allFields = classType.getDeclaredFields();
        for (Field field : allFields) {
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            // 只对含有@ExcelProperty注解的字段进行赋值
            if (excelProperty == null) {
                continue;
            }
            ExcelEntityField eef = new ExcelEntityField();

            String key = excelProperty.value().trim();// Excel Header名
            boolean required = excelProperty.required(); // 该列是否为必须列

            int index = _indexOfHeader(key);
            // 如果字段必须，而索引为-1 ，说明没有这一列，抛错
            if (required && index == -1) {
                throw new ExcelParseException("请填写" + key + "!");
            }

            eef.setField(field);
            eef.setColumnName(key);
            eef.setRequired(required);
            eef.setIndex(_indexOfHeader(key));
            eef.setAnnotation(excelProperty);

            eefs.add(eef);
        }
        return eefs;
    }


    /**
     * 获取字段的值，路由不同的字段类型
     *
     * @param value
     * @param eef
     * @return
     * @throws ExcelParseException
     * @throws InstantiationException
     * @throws IllegalAccessException
     * @throws ExcelContentInvalidException
     * @throws ExcelRegexpValidFailedException
     */
    @SuppressWarnings("rawtypes")
    private Object _getFieldValue(String value, ExcelEntityField eef) throws ExcelParseException, InstantiationException, IllegalAccessException, ExcelContentInvalidException, ExcelRegexpValidFailedException {
        // 进行规则校验
        ExcelProperty annotation = eef.getAnnotation();
        Class<? extends ExcelRule> rule = annotation.rule();

        // 获取解析后的字段结果
        Object result = null;
        try {
            // 是否提交过来的是空值
            // 如果提交值是空值而且含有默认值的话
            // 则让提交过来的空值为默认值
            if (("".equals(value) || (value == null)) && annotation.hasDefaultValue()) {
                value = annotation.defaultValue();
            }
            result = _parseFieldValue(value, eef.getField(), annotation.regexp());
        } catch (ExcelRegexpValidFailedException e) {
            // 捕获正则验证失败异常
            String errMsg = annotation.regexpErrorMessage();
            if ("".equals(errMsg)) {
                errMsg = "列 " + eef.getColumnName() + " 没有通过规则验证!";
            }
            throw new ExcelContentInvalidException(errMsg, e);
        } catch (NumberFormatException e) {
            throw new ExcelContentInvalidException("列 " + eef.getColumnName() + " 数据类型错误!");
        } catch (NullPointerException e) {
            throw new ExcelContentInvalidException("列 " + eef.getColumnName() + " 不能为空!");
        }
        /**
         * 缓存已经实例化过的规则对象，避免每次都重新
         * 创建新的对象的额外消耗
         */
        ExcelRule ruleObj = null;
        if (rulesCache.containsKey(rule.getName())) {
            ruleObj = rulesCache.get(rule.getName());
        } else {
            ruleObj = rule.newInstance();
            rulesCache.put(rule.getName(), ruleObj);
        }

        // 进行校验
        ruleObj.check(result, eef.getColumnName(), eef.getField().getName());

        return result;
    }

    /**
     * 解析字段类型
     *
     * @param value
     * @param field
     * @param regexp
     * @return
     * @throws ExcelParseException
     * @throws ExcelContentInvalidException
     * @throws ExcelRegexpValidFailedException
     */
    @SuppressWarnings("rawtypes")
    private Object _parseFieldValue(String value, Field field, String regexp) throws ExcelParseException, ExcelContentInvalidException, ExcelRegexpValidFailedException {
        Class<?> type = field.getType();
        String typeName = type.getName();
        // 字符串
        if ("java.lang.String".equals(typeName)) {
            if (!"".equals(regexp) && !value.matches(regexp)) {
                throw new ExcelRegexpValidFailedException("字段:" + field.getName() + "验证失败");
            }
            return value;
        }
        // 长整形
        if ("java.lang.Long".equals(typeName) || "long".equals(typeName)) {
            if (!"".equals(regexp) && !value.matches(regexp)) {
                throw new ExcelRegexpValidFailedException("字段:" + field.getName() + "验证失败");
            }
            return Long.parseLong(value);
        }
        // 整形
        if ("java.lang.Integer".equals(typeName) || "int".equals(typeName)) {
            if (!"".equals(regexp) && !value.matches(regexp)) {
                throw new ExcelRegexpValidFailedException("字段:" + field.getName() + "验证失败");
            }
            return Integer.parseInt(value);
        }
        // 短整型
        if ("java.lang.Short".equals(typeName) || "short".equals(typeName)) {
            if (!"".equals(regexp) && !value.matches(regexp)) {
                throw new ExcelRegexpValidFailedException("字段:" + field.getName() + "验证失败");
            }
            return Short.parseShort(value);
        }
        // Date型
        if ("java.util.Date".equals(typeName)) {
            try {
                return new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").parse(value);
            } catch (ParseException e) {
                throw new ExcelParseException("日期类型格式有误!");
            }
        }
        // Timestamp
        if ("java.sql.Timestamp".equals(typeName)) {
            try {
                return new Timestamp(new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").parse(value).getTime());
            } catch (ParseException e) {
                throw new ExcelParseException("时间戳类型格式有误!");
            }
        }
        // Char型
        if ("java.lang.Character".equals(typeName) || "char".equals(typeName)) {
            if (value.length() == 1) {
                return value.charAt(0);
            }
        }
        // 用户注册的自定义类型
        for (Class<? extends ExcelFieldType> et : userDefinedType) {
            if (et.getName().equals(typeName)) {
                try {
                    ExcelFieldType newInstance = et.newInstance();
                    return newInstance.parseValue(value);
                } catch (InstantiationException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
        throw new ExcelParseException("不支持的字段类型 " + typeName + " !");
    }


}
