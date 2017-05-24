package com.geese.plugin.excel;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Collection;
import java.util.UUID;

/**
 * Created by Administrator on 2017/5/22.
 */
public class ExcelValidation {
    // 验证范围
    private Integer firstRow;
    private Integer lastRow;
    private Integer firstCol;
    private Integer lastCol;
    // 约束规则
    private Integer validationType;
    private Integer operatorType;
    // 约束值
    private Collection constraintValues;
    // 显示下拉箭头
    private Boolean suppressDropDownArrow;
    // 显示错误提示框
    private Boolean showErrorBox;

    public ExcelValidation() {
        this.suppressDropDownArrow = true;
        this.showErrorBox = true;
    }

    public ExcelValidation(Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol, Collection constraintValues) {
        this();
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;
        this.constraintValues = constraintValues;
    }

    public ExcelValidation(Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol, Integer validationType, Integer operatorType, Collection constraintValues) {
        this(firstRow, lastRow, firstCol, lastCol, constraintValues);
        this.validationType = validationType;
        this.operatorType = operatorType;
    }

    public Integer getFirstRow() {
        return firstRow;
    }

    public void setFirstRow(Integer firstRow) {
        this.firstRow = firstRow;
    }

    public Integer getLastRow() {
        return lastRow;
    }

    public void setLastRow(Integer lastRow) {
        this.lastRow = lastRow;
    }

    public Integer getFirstCol() {
        return firstCol;
    }

    public void setFirstCol(Integer firstCol) {
        this.firstCol = firstCol;
    }

    public Integer getLastCol() {
        return lastCol;
    }

    public void setLastCol(Integer lastCol) {
        this.lastCol = lastCol;
    }

    public Integer getValidationType() {
        return validationType;
    }

    public void setValidationType(Integer validationType) {
        this.validationType = validationType;
    }

    public Integer getOperatorType() {
        return operatorType;
    }

    public void setOperatorType(Integer operatorType) {
        this.operatorType = operatorType;
    }

    public Collection getConstraintValues() {
        return constraintValues;
    }

    public void setConstraintValues(Collection constraintValues) {
        this.constraintValues = constraintValues;
    }

    public Boolean getSuppressDropDownArrow() {
        return suppressDropDownArrow;
    }

    public void setSuppressDropDownArrow(Boolean suppressDropDownArrow) {
        this.suppressDropDownArrow = suppressDropDownArrow;
    }

    public Boolean getShowErrorBox() {
        return showErrorBox;
    }

    public void setShowErrorBox(Boolean showErrorBox) {
        this.showErrorBox = showErrorBox;
    }

    // 把校验添加到sheet中 -------------------------------------
    public void addToSheet(Sheet sheet) {
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidation validation = null;
        String[] values = (String[]) constraintValues.toArray(new String[constraintValues.size()]);

        boolean isListValidationType = (null == validationType);
        String listFormula = null;

        if (isListValidationType) {
            Workbook wk = sheet.getWorkbook();
            String hiddenSheetName = "__Hidden_Sheet__";
            Sheet hiddenSheet = wk.getSheet(hiddenSheetName);
            if (null == hiddenSheet) {
                hiddenSheet = wk.createSheet(hiddenSheetName);
                wk.setSheetHidden(wk.getSheetIndex(hiddenSheet), true);
            }
            // 获取没有设置值的单元格
            Row firstRow = ExcelHelper.createRow(hiddenSheet, 0);
            Cell cell = null;
            for (int i = 0; i < 16384; i++) {
                cell = firstRow.getCell(i);
                if (null == cell) {
                    cell = firstRow.createCell(i);
                    break;
                }
            }
            // 单元格列对应的字母
            int columnIndex = cell.getColumnIndex();
            String colAlphabet = CellReference.convertNumToColString(columnIndex);
            // 设置下拉框的值
            for (int i = 0; i < values.length; i++) {
                ExcelHelper.createRow(hiddenSheet, i).createCell(columnIndex).setCellValue(values[i]);
            }
            // 给下拉框设置一个名称
            Name namedCell = wk.createName();
            listFormula = hiddenSheetName + colAlphabet + "1" + colAlphabet + values.length;
            namedCell.setNameName(listFormula);
            namedCell.setRefersToFormula(hiddenSheetName + "!$" + colAlphabet + "$1:$" + colAlphabet + "$" + values.length);
        }

        // 判断sheet类型，根据sheet类型使用对应的API来创建 Data Validation Constraint
        if (sheet instanceof HSSFSheet) {
            DVConstraint dvConstraint = null;
            if (isListValidationType) {
                dvConstraint = DVConstraint.createFormulaListConstraint(listFormula);
            } else {
                switch (validationType) {
                    case DataValidationConstraint.ValidationType.INTEGER:
                    case DataValidationConstraint.ValidationType.DECIMAL:
                    case DataValidationConstraint.ValidationType.TEXT_LENGTH:
                        dvConstraint = DVConstraint.createNumericConstraint(validationType, operatorType, values[0], values.length == 2 ? values[1] : null);
                }
            }
            validation = new HSSFDataValidation(addressList, dvConstraint);
        } else {
            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
            DataValidationConstraint dvConstraint = null;
            if (isListValidationType) {
                dvConstraint = dvHelper.createFormulaListConstraint(listFormula);
            } else {
                switch (validationType) {
                    case DataValidationConstraint.ValidationType.INTEGER:
                    case DataValidationConstraint.ValidationType.DECIMAL:
                    case DataValidationConstraint.ValidationType.TEXT_LENGTH:
                        dvConstraint = dvHelper.createNumericConstraint(validationType, operatorType, values[0], values.length == 2 ? values[1] : null);
                }
            }
            validation = dvHelper.createValidation(dvConstraint, addressList);
        }
        // 设置下拉箭头
        validation.setSuppressDropDownArrow(suppressDropDownArrow);
        // 显示错误提示
        validation.setShowErrorBox(showErrorBox);
        sheet.addValidationData(validation);
    }
}
