package com.geese.plugin.excel.config;

import com.geese.plugin.excel.filter.*;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * sheet中数据表格的配置信息
 *
 * @author zhangguangyong <a href="#">1243610991@qq.com</a>
 * @date 2016/11/16 16:04
 * @sine 0.0.1
 */
public class Table extends Filterable {
    /**
     * 映射sheet中一个列表的头部单元格
     */
    private List<Point> columns;

    /**
     * 读取sheet的开始行
     */
    private Integer startRow;

    /**
     * 读取sheet的行数
     */
    private Integer endRow;

    /**
     * table的数据源
     */
    private List data;

    /**
     * table配置关联的sheet配置信息
     */
    private Sheet sheet;

    private FilterChain rowBeforeReadFilterChain;

    private FilterChain rowAfterReadFilterChain;

    private FilterChain rowBeforeWriteFilterChain;

    private FilterChain rowAfterWriteFilterChain;

    private FilterChain cellBeforeReadFilterChain;

    private FilterChain cellAfterReadFilterChain;

    private FilterChain cellBeforeWriteFilterChain;

    private FilterChain cellAfterWriteFilterChain;

    public Table() {
        this.startRow = 0;
    }


    public Table addColumn(Point column) {
        if (null == this.columns) {
            this.columns = new ArrayList<>();
        }
        this.columns.add(column);
        return this;
    }

    public Table addColumns(Collection<Point> columns) {
        if (null == this.columns) {
            this.columns = new ArrayList<>();
        }
        this.columns.addAll(columns);
        return this;
    }

    public List<Point> getColumns() {
        return columns;
    }

    public void setColumns(List<Point> columns) {
        this.columns = columns;
    }

    public Integer getStartRow() {
        return startRow;
    }

    public void setStartRow(Integer startRow) {
        this.startRow = startRow;
    }

    public Integer getEndRow() {
        return endRow;
    }

    public void setEndRow(Integer endRow) {
        this.endRow = endRow;
    }

    public List getData() {
        return data;
    }

    public void setData(List data) {
        this.data = data;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public void addRowBeforeReadFilter(RowBeforeReadFilter filter) {
        if (null == rowBeforeReadFilterChain) {
            rowBeforeReadFilterChain = new FilterChain();
        }
        rowBeforeReadFilterChain.addFilter(filter);
    }

    public void addRowBeforeReadFilters(Collection<RowBeforeReadFilter> filters) {
        if (null == rowBeforeReadFilterChain) {
            rowBeforeReadFilterChain = new FilterChain();
        }
        rowBeforeReadFilterChain.addFilters(filters);
    }

    public void addRowAfterReadFilter(RowAfterReadFilter filter) {
        if (null == rowAfterReadFilterChain) {
            rowAfterReadFilterChain = new FilterChain();
        }
        rowAfterReadFilterChain.addFilter(filter);
    }

    public void addRowAfterReadFilters(Collection<RowAfterReadFilter> filters) {
        if (null == rowAfterReadFilterChain) {
            rowAfterReadFilterChain = new FilterChain();
        }
        rowAfterReadFilterChain.addFilters(filters);
    }

    public void addCellBeforeReadFilter(CellBeforeReadFilter filter) {
        if (null == cellBeforeReadFilterChain) {
            cellBeforeReadFilterChain = new FilterChain();
        }
        cellBeforeReadFilterChain.addFilter(filter);
    }

    public void addCellBeforeReadFilters(Collection<CellBeforeReadFilter> filters) {
        if (null == cellBeforeReadFilterChain) {
            cellBeforeReadFilterChain = new FilterChain();
        }
        cellBeforeReadFilterChain.addFilters(filters);
    }

    public void addCellAfterReadFilter(CellAfterReadFilter filter) {
        if (null == cellAfterReadFilterChain) {
            cellAfterReadFilterChain = new FilterChain();
        }
        cellAfterReadFilterChain.addFilter(filter);
    }

    public void addCellAfterReadFilters(Collection<CellAfterReadFilter> filters) {
        if (null == cellAfterReadFilterChain) {
            cellAfterReadFilterChain = new FilterChain();
        }
        cellAfterReadFilterChain.addFilters(filters);
    }

    public void addRowBeforeWriteFilter(RowBeforeWriteFilter filter) {
        if (null == rowBeforeWriteFilterChain) {
            rowBeforeWriteFilterChain = new FilterChain();
        }
        rowBeforeWriteFilterChain.addFilter(filter);
    }

    public void addRowBeforeWriteFilters(Collection<RowBeforeWriteFilter> filters) {
        if (null == rowBeforeWriteFilterChain) {
            rowBeforeWriteFilterChain = new FilterChain();
        }
        rowBeforeWriteFilterChain.addFilters(filters);
    }

    public void addRowAfterWriteFilter(RowAfterWriteFilter filter) {
        if (null == rowAfterWriteFilterChain) {
            rowAfterWriteFilterChain = new FilterChain();
        }
        rowAfterWriteFilterChain.addFilter(filter);
    }

    public void addRowAfterWriteFilters(Collection<RowAfterWriteFilter> filters) {
        if (null == rowAfterWriteFilterChain) {
            rowAfterWriteFilterChain = new FilterChain();
        }
        rowAfterWriteFilterChain.addFilters(filters);
    }

    public void addCellBeforeWriteFilter(CellBeforeWriteFilter filter) {
        if (null == cellBeforeWriteFilterChain) {
            cellBeforeWriteFilterChain = new FilterChain();
        }
        cellBeforeWriteFilterChain.addFilter(filter);
    }

    public void addCellBeforeWriteFilters(Collection<CellBeforeWriteFilter> filters) {
        if (null == cellBeforeWriteFilterChain) {
            cellBeforeWriteFilterChain = new FilterChain();
        }
        cellBeforeWriteFilterChain.addFilters(filters);
    }

    public void addCellAfterWriteFilter(CellAfterWriteFilter filter) {
        if (null == cellAfterWriteFilterChain) {
            cellAfterWriteFilterChain = new FilterChain();
        }
        cellAfterWriteFilterChain.addFilter(filter);
    }

    public void addCellAfterWriteFilters(Collection<CellAfterWriteFilter> filters) {
        if (null == cellAfterWriteFilterChain) {
            cellAfterWriteFilterChain = new FilterChain();
        }
        cellAfterWriteFilterChain.addFilters(filters);
    }
}
