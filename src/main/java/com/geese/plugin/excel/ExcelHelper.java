package com.geese.plugin.excel;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import javax.script.Bindings;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * ExcelMapping 辅助类
 * <p>在对Excel进行read或write操作的时候，所需要用到的一些辅助接口</p>
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:09
 * @sine 0.0.1
 */
public class ExcelHelper {
    /**
     * 脚本引用，在对excel进行read操作的时候，执行where条件
     */
    private static ScriptEngine engine = new ScriptEngineManager().getEngineByName("javascript");

    /**
     * like 关键字匹配，用于where条件中like关键字的匹配
     */
    private static Pattern likePattern = Pattern.compile("\\s*[^\\s]+\\s+(not)?like\\s+[^\\s)]+");

    /**
     * in 关键字匹配，用于where条件中in关键字的匹配
     */
    private static Pattern inPattern = Pattern.compile("\\s*[^\\s]+\\s+(not)?in\\s+[^\\s)]+");

    /**
     * 查询(read)关键字 from where limit
     */
    private static String[] queryKeys = {OperationKey.FROM.name(), OperationKey.WHERE.name(), OperationKey.LIMIT.name()};

    /**
     * 插入(write)关键字 into limit
     */
    private static String[] insertKeys = {OperationKey.INTO.name(), OperationKey.LIMIT.name()};


    /**
     * 断言Excel是xls类型（2003）
     *
     * @param input
     * @return
     */
    public static boolean isXls(InputStream input) {
        try {
            new XSSFWorkbook(input);
            return false;
        } catch (IOException e) {
            // e.printStackTrace();
            return true;
        }
    }

    /**
     * 断言Excel是Xlsx类型（2007）
     *
     * @param input
     * @return
     */
    public static boolean isXlsx(InputStream input) {
        return !isXls(input);
    }

    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public static Object getCellValue(Cell cell) {
        if (null == cell) {
            return null;
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue().getString();
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_FORMULA:
//                return cell.getCellFormula();
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_BLANK:
                return null;
            default:
                return null;
        }
    }

    /**
     * 设置单元格的值
     *
     * @param cell
     * @param value
     * @return
     */
    public static Cell setCellValue(Cell cell, Object value) {
        if (null == cell || null == value) {
            return cell;
        }
        Class<?> valueClass = value.getClass();
        // String
        if (String.class.isAssignableFrom(valueClass)) {
            cell.setCellValue((String) value);
            return cell;
        }
        // Boolean
        if (Boolean.class.isAssignableFrom(valueClass)) {
            cell.setCellValue((Boolean) value);
            return cell;
        }
        // Date
        if (Date.class.isAssignableFrom(valueClass)) {
            cell.setCellValue((Date) value);
            return cell;
        }
        // Calendar
        if (Calendar.class.isAssignableFrom(valueClass)) {
            cell.setCellValue((Calendar) value);
            return cell;
        }
        // Double
        if (Double.class.isAssignableFrom(valueClass)) {
            cell.setCellValue((Double) value);
            return cell;
        }
        // RichTextString
        if (RichTextString.class.isAssignableFrom(valueClass)) {
            cell.setCellValue((RichTextString) value);
            return cell;
        }
        // 剩余类型都当做String来处理
        cell.setCellValue(value.toString());
        return cell;
    }

    /**
     * 创建单元格（存在：直接返回，不存在：创建）
     *
     * @param row
     * @param cellIndex
     * @return
     */
    public static Cell createCell(Row row, int cellIndex) {
        Cell cell = row.getCell(cellIndex);
        if (null != cell) {
            return cell;
        }
        return row.createCell(cellIndex);
    }

    /**
     * 创建行（存在：直接返回，不存在：创建）
     *
     * @param sheet
     * @param rowIndex
     * @return
     */
    public static Row createRow(org.apache.poi.ss.usermodel.Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (null != row) {
            return row;
        }
        return sheet.createRow(rowIndex);
    }

    /**
     * 查询条件过滤
     *
     * @param where
     * @param conditionMap
     * @param parameterMap
     * @return
     * @throws ScriptException
     */
    public static boolean whereFilter(String where, Map conditionMap, Map parameterMap) {
        boolean hasParameter = where.contains(":") || where.contains("?");
        Bindings bindings = engine.createBindings();
        String parsedWhere = where;
        // 名称的参数和占位符处理
        if (hasParameter) {
            parsedWhere = bindVariableToParameterizedScript(where, parameterMap, bindings);
        }
        bindings.putAll(conditionMap);
        // 逻辑运算符转义
        parsedWhere = symbolEscape(parsedWhere);
        // like 操作解析
        parsedWhere = likeOperationParse(parsedWhere, bindings);
        // in 操作解析
        parsedWhere = inOperationParse(parsedWhere, bindings);
        // 执行脚本
        try {
            System.out.println("解析后的where:" + parsedWhere + ", 绑定的参数：" + new LinkedHashMap(bindings));

            return true == Boolean.valueOf(engine.eval(parsedWhere, bindings).toString());
        } catch (ScriptException e) {
            // TODO where 脚本解析异常处理
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 绑定变量到 -> 参数化脚本
     *
     * @param parameterizedScript
     * @param variableMap
     * @param bindings
     * @return
     */
    private static String bindVariableToParameterizedScript(String parameterizedScript, Map variableMap, Bindings bindings) {
        String script = String.valueOf(parameterizedScript);
        String uuid = "_" + UUID.randomUUID().toString().replace("-", "");
        // 命名的参数
        if (script.contains(":")) {
            script = script.replaceAll("\\:", uuid);
            Set variableSet = variableMap.keySet();
            for (Object variable : variableSet) {
                bindings.put(uuid + variable, variableMap.get(variable));
            }
        }
        // 占位符参数
        else if (script.contains("?")) {
            Collection values = variableMap.values();
            int index = 0;
            for (Object value : values) {
                String variable = uuid + "_" + (index++);
                script = script.replaceFirst("\\?", variable);
                bindings.put(variable, value);
            }
        }
        return script;
    }

    /**
     * 符号转义 [and -> &&] [or -> ||]
     *
     * @param str
     * @return
     */
    private static String symbolEscape(String str) {
        String ref = String.valueOf(str);
        ref = ref.replaceAll("and\\s+", "&& ");
        ref = ref.replaceAll("and\\(", "&&(");
        ref = ref.replaceAll("or\\s+", "|| ");
        ref = ref.replaceAll("or\\(", "||(");
        return ref;
    }

    /**
     * like 关键字操作解析
     *
     * @param where
     * @return
     */
    private static String likeOperationParse(String where, Map parameterMap) {
        // like -> value.test() 比如：name like /.*zhangsan.*/i -> /.*zhangsan.*/i.test(name)
        String ref = String.valueOf(where);
        Matcher m = likePattern.matcher(ref);
        while (m.find()) {
            String matched = m.group().trim();
            // 拆开 -> 处理 -> 重新组装
            String[] likeCondition = matched.split("\\s+");
            String variable = likeCondition[0];
            String operation = likeCondition[1];
            String parameterVariable = likeCondition[2];
            // 可能是 %xxx% 或者 /xxx/
            String parameterValue = parameterMap.get(parameterVariable).toString();
            // [like %name -> like /.*/i], [like %name% -> like /.*name.*/i], [like _name -> like /.name/i]
            if (!parameterValue.matches("\\/.+\\/.?")) {
                parameterValue = parameterValue.replaceAll("%", ".*").replaceAll("_", ".");
                parameterValue = "/^" + parameterValue + "$/i";
            }
            // like:/xxx/ -> notlike:!/xxx/
            if (operation.startsWith("not")) {
                parameterValue = "!" + parameterValue;
            }
            // 组装成：/xxx/.test(variable)
            String regex = parameterValue + ".test(" + variable + ")";
            ref = ref.replace(matched, regex);
        }
        return ref;
    }

    /**
     * in 关键字操作解析
     *
     * @param where
     * @param parameterMap
     * @return
     */
    private static String inOperationParse(String where, Map parameterMap) {
        String ref = String.valueOf(where);
        Matcher m = inPattern.matcher(ref);
        while (m.find()) {
            // 匹配到的
            String matched = m.group().trim();
            // 拆开 -> 处理 -> 重新组装
            // 拆开
            String[] likeCondition = matched.split("\\s+");
            String variable = likeCondition[0];
            String operation = likeCondition[1];
            String parameterVariable = likeCondition[2];
            Object parameterValue = parameterMap.get(parameterVariable);
            // 处理 name in xxx -> /^x1$|^x2$|^x3$/.test(name)
            Iterator parameterValueIterator;
            // 集合
            if (Iterable.class.isAssignableFrom(parameterValue.getClass())) {
                parameterValueIterator = ((Iterable) parameterValue).iterator();
            }
            // 数组和单值
            else {
                parameterValueIterator = Arrays.asList(parameterValue).iterator();
            }
            StringBuffer inValueRegex = new StringBuffer();
            while (parameterValueIterator.hasNext()) {
                inValueRegex.append("^" + parameterValueIterator.next() + "$|");
            }
            inValueRegex = inValueRegex.deleteCharAt(inValueRegex.length() - 1);
            String regex = "/" + inValueRegex.toString() + "/";
            // in:/xxx/ -> notin:!/xxx/
            if (operation.startsWith("not")) {
                regex = "!" + regex;
            }
            // 重新组装
            regex = regex + ".test(" + variable + ")";
            ref = ref.replace(matched, regex);
        }
        return ref;
    }

    /**
     * 查询语句关键字解析
     *
     * @param query
     * @return
     */
    public static Map<OperationKey, String> parseQuery(String query) {
        return parseKeys(query, queryKeys);
    }

    /**
     * 插入语句关键字解析
     *
     * @param insert
     * @return
     */
    public static Map<OperationKey, String> parseInsert(String insert) {
        return parseKeys(insert, insertKeys);
    }

    /**
     * 关键字解析
     * 1. from -> `FROM`, where -> `WHERE`
     * 2. 1 name from sheet where name like % limit 10 30 -> [1 name, name like %, 10 30]
     *
     * @param query
     * @param keys
     * @return
     */
    private static Map<OperationKey, String> parseKeys(String query, String[] keys) {
        // 删除多余的空格
        String parsedQuery = String.valueOf(query).replaceAll("\\s+", " ");
        // 将`key` 转换为大写 -> `KEY`
        Pattern p = Pattern.compile("`[^`]+`");
        Matcher m = p.matcher(parsedQuery);
        String matched;
        while (m.find()) {
            matched = m.group();
            parsedQuery = parsedQuery.replace(matched, matched.toUpperCase());
        }
        // 查询语句使用的key
        List<String> usedKeys = new ArrayList<>();
        // 将key转换为标准的key, from -> `FROM`, where -> `WHERE`
        for (String key : keys) {
            // 获取一个标准格式的key
            String wrapKey = wrapKey(key);
            // 如果query没有使用标准格式的key，则转换为标准的key格式
            if (!parsedQuery.contains(wrapKey)) {
                p = Pattern.compile("\\s+" + key + "\\s+", Pattern.CASE_INSENSITIVE);
                m = p.matcher(query);
                if (m.find()) {
                    matched = m.group();
                    parsedQuery = parsedQuery.replace(matched, " " + wrapKey + " ");
                    usedKeys.add(wrapKey);
                }
            } else {
                usedKeys.add(wrapKey);
            }
        }
        // 获取key对应的数据 1 name from sheet where name like % limit 10 30 -> [1 name, name like %, 10 30]
        Map<OperationKey, String> keyDataMap = new LinkedHashMap();
        int nextKeyIndex = 0;
        int index = 0;
        for (String usedKey : usedKeys) {
            int keyIndex = parsedQuery.indexOf(usedKey);
            String keyData = parsedQuery.substring(nextKeyIndex, keyIndex).trim();
            // 不管什么查询第一个都需要指定查询的 column
            keyDataMap.put(index == 0 ? OperationKey.COLUMN : OperationKey.valueOf(unWrapKey(usedKeys.get(index - 1))), keyData);
            nextKeyIndex = keyIndex + usedKey.length();
            index++;
        }
        // 最有一个关键字的数据
        keyDataMap.put(OperationKey.valueOf(unWrapKey(usedKeys.get(usedKeys.size() - 1))), parsedQuery.substring(nextKeyIndex).trim());
        return keyDataMap;
    }

    /**
     * 包装关键字
     *
     * @param key
     * @return
     */
    public static String wrapKey(String key) {
        return "`" + key + "`";
    }

    /**
     * 解包装关键字
     *
     * @param key
     * @return
     */
    public static String unWrapKey(String key) {
        return key.replace("`", "");
    }

    public static boolean isNumber(String text) {
        return null != text && text.matches("([0-9])|([1-9]\\d*)");
    }

    /**
     * 获取Sheet中的图片
     *
     * @param sheet
     * @return key: row_column
     */
    public static Map<String, PictureData> getPictures(Sheet sheet) {
        if (sheet instanceof HSSFSheet) {
            return getSheetPictures((HSSFSheet) sheet);
        }
        return getXSheetPictures((XSSFSheet) sheet);
    }

    /**
     * 获取Row中的图片
     *
     * @param row
     * @return key: column
     */
    public static Map<Integer, PictureData> getPictures(Row row) {
        if (row instanceof HSSFRow) {
            return getSheetPictures((HSSFRow) row);
        }
        return getXSheetPictures((XSSFRow) row);
    }

    /**
     * 为Cell设置图片
     *
     * @param cell
     * @param pictureData
     * @return
     */
    public static Picture setPicture(Cell cell, byte[] pictureData) {
        if (null == pictureData || pictureData.length <= 0) {
            return null;
        }
        Sheet sheet = cell.getSheet();
        Workbook workbook = sheet.getWorkbook();
        CreationHelper helper = workbook.getCreationHelper();

        // 添加图片到Workbook
        int pictureIndex = workbook.addPicture(pictureData, Workbook.PICTURE_TYPE_JPEG);

        // 画板，把图片放入某个位置
        Drawing drawing = sheet.getDrawingPatriarch();
        if (null == drawing) {
            drawing = sheet.createDrawingPatriarch();
        }
        // 苗点，确定图片位置
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setRow1(cell.getRowIndex());
        Picture picture = drawing.createPicture(anchor, pictureIndex);
        picture.resize(1.0, 1.0);

        return picture;
    }


    /**
     * 获取 HSSFSheet中的图片
     *
     * @param sheet
     * @return
     */
    private static Map<String, PictureData> getSheetPictures(HSSFSheet sheet) {
        HSSFWorkbook workbook = sheet.getWorkbook();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (pictures.isEmpty()) {
            return null;
        }
        Map<String, PictureData> sheetIndexPicMap = new LinkedHashMap<String, PictureData>();
        for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
            HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
            if (shape instanceof HSSFPicture) {
                HSSFPicture pic = (HSSFPicture) shape;
                int pictureIndex = pic.getPictureIndex() - 1;
                HSSFPictureData picData = pictures.get(pictureIndex);
                String picIndex = anchor.getRow1() + "_" + anchor.getCol1();
                sheetIndexPicMap.put(picIndex, picData);
            }
        }
        return sheetIndexPicMap;
    }

    private static Map<Integer, PictureData> getSheetPictures(HSSFRow row) {
        HSSFSheet sheet = row.getSheet();
        HSSFWorkbook workbook = sheet.getWorkbook();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (pictures.isEmpty()) {
            return null;
        }

        int resultRowNum = row.getRowNum();
        Map<Integer, PictureData> sheetIndexPicMap = new LinkedHashMap<Integer, PictureData>();
        for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
            HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
            if (shape instanceof HSSFPicture && resultRowNum == anchor.getRow1()) {
                HSSFPicture pic = (HSSFPicture) shape;
                int pictureIndex = pic.getPictureIndex() - 1;
                HSSFPictureData picData = pictures.get(pictureIndex);
                sheetIndexPicMap.put(Integer.valueOf(anchor.getCol1()), picData);
            }
        }
        return sheetIndexPicMap;
    }

    /**
     * 获取 XSSFSheet 中的图片
     *
     * @param sheet
     * @return
     */
    private static Map<String, PictureData> getXSheetPictures(XSSFSheet sheet) {
        List<POIXMLDocumentPart> relations = sheet.getRelations();
        if (relations.isEmpty()) {
            return null;
        }

        Map<String, PictureData> sheetIndexPicMap = new LinkedHashMap<String, PictureData>();
        for (POIXMLDocumentPart documentPart : relations) {
            if (documentPart instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) documentPart;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    int rowNum = ctMarker.getRow();

                    String picIndex = rowNum + "_" + ctMarker.getCol();
                    sheetIndexPicMap.put(picIndex, pic.getPictureData());
                }
            }
        }
        return sheetIndexPicMap;
    }

    private static Map<Integer, PictureData> getXSheetPictures(XSSFRow row) {
        XSSFSheet sheet = row.getSheet();
        List<POIXMLDocumentPart> relations = sheet.getRelations();
        if (relations.isEmpty()) {
            return null;
        }

        int resultRowNum = row.getRowNum();
        Map<Integer, PictureData> sheetIndexPicMap = new LinkedHashMap<Integer, PictureData>();
        for (POIXMLDocumentPart documentPart : relations) {
            if (documentPart instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) documentPart;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    int rowNum = ctMarker.getRow();
                    if (rowNum == resultRowNum) {
                        sheetIndexPicMap.put(ctMarker.getCol(), pic.getPictureData());
                    }
                }
            }
        }
        return sheetIndexPicMap;
    }

}