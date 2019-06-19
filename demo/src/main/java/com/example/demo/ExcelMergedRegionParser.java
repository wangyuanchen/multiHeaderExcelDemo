package com.example.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;


public class ExcelMergedRegionParser {
    // 阈值 临界值
    private static int THRESHOLD = 3;
    // 数据行起始位置，默认为-1没有符合条件的数据行
    private static int dataLineNum = -1;
    // Excel对象
    private static Workbook workbook = null;
    // sheet下标，默认为0
    private static int sheetNum = 0;

    /**
     *
     * @param workbook Excel对象
     */
    public ExcelMergedRegionParser(Workbook workbook){
        this.workbook = workbook;
    }
    /**
     *
     * @param workbook Excel对象
     * @param sheetNum sheet下标
     */
    public ExcelMergedRegionParser(Workbook workbook, int sheetNum){
        this.workbook = workbook;
        this.sheetNum = sheetNum;
    }

    /**
     *
     * @param workbook Excel对象
     * @param sheetNum sheet下标
     * @param THRESHOLD 临界值
     */
    public ExcelMergedRegionParser(Workbook workbook, int sheetNum, int THRESHOLD){
        this.workbook = workbook;
        this.sheetNum = sheetNum;
        this.THRESHOLD = THRESHOLD;
    }


    /**
     *
     * @return 返回解析后的EXCEL最后一行汉字标题，有序
     */
    public List<String> getHeaderList(){
        List<String> headList = new ArrayList<>();
            // 获取sheet页签
            Sheet sheet = workbook.getSheetAt(sheetNum);
            // 获取工具类实例
            ExcelMergedRegionUtil excelMergedRegionUtil = ExcelMergedRegionUtil.getInstance();
            // 该sheet中是否含有合并单元格
            if (excelMergedRegionUtil.hasMerged(sheet)){
                    int rows = sheet.getPhysicalNumberOfRows();
                    for (int i = 0;i<rows;i++){
                        int count = 0;
                        Row row = sheet.getRow(i);
                        if(row == null)
                            continue;
                       int cells = row.getPhysicalNumberOfCells();
                       for(int j = 0 ;j<cells;j++){
                            Cell cell = row.getCell(j);
                            // 如果有一列为null，则跳过
                            if (cell == null)
                                continue;
                            // 判断是数字类型就加1
                            if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
                                count++;
                        }
                        // 如果计数器大于阈值说明找到了第一个数据行
                        if (count >= THRESHOLD){
                            // 数据行所在位置
                            dataLineNum = i;
                            row = sheet.getRow(i-1);
                            // 注释的代码，是为了防止阈值为0时报NULLpoint异常
                            if (row == null)
                                continue;
                            cells = row.getPhysicalNumberOfCells();
                            for(int j = 0 ;j<cells;j++){
                                Cell cell = row.getCell(j);
                                // 判断如果有空就跳过，这里加1是为了弥补损失
                                if (cell == null){
                                    cells++;
                                    continue;
                                }
                                // 判断每一个单元格是否为合并单元格
                                boolean flag = excelMergedRegionUtil.isMergedRegion(sheet,i-1,j);
                                String value = "";
                                // true
                                if(flag){
                                    // 判断该单元格为""并且是行合并才跳过
                                    if("".equals(excelMergedRegionUtil.getCellValue(cell))&&excelMergedRegionUtil.isMergedRow(sheet,i-1,j))
                                        continue;
                                    // 获取合并单元格的值
                                    value = excelMergedRegionUtil.getMergedRegionValue(sheet,i-1,j);
                                    headList.add(value);
                                }else {
                                    value = excelMergedRegionUtil.getCellValue(cell);
                                    headList.add(value);
                                }
                            }
                            break;
                        }
                    }
            }else{
                // 获取最大物理行行数
                int rows = sheet.getPhysicalNumberOfRows();
                for (int i = 0;i<rows;i++){
                    // 数值类型计数器
                    int count = 0;
                    Row row = sheet.getRow(i);
                    if(row == null)
                        continue;
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()){
                        Cell cell = cellIterator.next();
                        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
                            // 每次加1
                            count++;
                    }
                    // 如果计数器大于阈值说明找到了第一个数据行
                    if (count >= THRESHOLD){
                        // 数据行位置
                        dataLineNum = i;
                        // 获取标题行位置
                        row = sheet.getRow(i-1);
                        // 注释的代码，是为了防止阈值为0时报NULLpoint异常
                        if (row == null)
                            continue;
                        Iterator<Cell> cellSpecIterator = row.cellIterator();
                        while (cellSpecIterator.hasNext()){
                            Cell cell = cellSpecIterator.next();
                            String value = excelMergedRegionUtil.getCellValue(cell);
                            headList.add(value);
                        }
                        break;
                    }
                }
            }
        return headList;
    }

    /**
     *
     * @return 数据行下标
     */
    public int getDataLineNum(){
        return this.dataLineNum;
    }

    /**
     *  得到用户自定义表头布局
     * @return 返回递归得到的标题List
     */
    public List<Map<String,Object>> getHeaderLayout(){
        List<Map<String,Object>> headList = new ArrayList<>();
        // 获取sheet页签
        Sheet sheet = workbook.getSheetAt(sheetNum);
        // 获取最大物理行行数
        int rows = sheet.getPhysicalNumberOfRows();
        for (int i = 0;i<rows;i++){
                int count = 0;
                Row row = sheet.getRow(i);
                if(row == null)
                    continue;
                int cells = row.getPhysicalNumberOfCells();
                for(int j = 0 ;j<cells;j++){
                    Cell cell = row.getCell(j);
                    // 如果有一列为null，则跳过
                    if (cell == null)
                        continue;
                    // 判断是数字类型就加1
                    if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
                        count++;
                }
                // 如果计数器大于阈值说明找到了第一个数据行
                if (count >= THRESHOLD){
                    // 数据行所在位置
                    dataLineNum = i;
                    // 进行递归拿到标题List<Map<String,Object>>
                    headList = getCellBean(0,cells,0,0) ;
                    break;
                }
            }
        return headList;
    }

    /**
     *
     * @param firstColumn 循环开始列
     * @param lastColumn 循环结束咧
     * @param firstRow 循环开始行
     * @param lastRow 循环结束行
     * @return
     */
    public List<Map<String,Object>> getCellBean(int firstColumn, int lastColumn, int firstRow, int lastRow){
        // 递归结束条件 判断到达数据行结束
        if (lastRow==dataLineNum)
            return null;
        List<Map<String,Object>> headList = new ArrayList<>();
        // 获取sheet页签
        Sheet sheet = workbook.getSheetAt(sheetNum);
        // 获取工具类实例
        ExcelMergedRegionUtil excelMergedRegionUtil = ExcelMergedRegionUtil.getInstance();
       //  双重循环开始
        for (int i = firstRow;i<=lastRow;i++){
            Row row = sheet.getRow(i);
            for (int j = firstColumn;j<=lastColumn;j++){
                Cell cell = row.getCell(j);
                // 如果是合并单元格就不重复进行赋值和递归了，如果合并单元格的值为空，那么细粒度标题也不会出现在List中
                if (excelMergedRegionUtil.getCellValue(cell)==null||"".equals(excelMergedRegionUtil.getCellValue(cell)))
                    continue;
                Map<String,Object> map = new HashMap<>();
                // 判断该单元格是否为合并单元格
                if (excelMergedRegionUtil.isMergedRegion(sheet,i,j)){
                    // 获取合并单元格的值
                    map.put("value",excelMergedRegionUtil.getMergedRegionValue(sheet,i,j));
                }else {
                    // 根据类型去单个单元格的值
                    map.put("value",excelMergedRegionUtil.getCellValue(cell));
                }
                int firstCol =excelMergedRegionUtil.getMergedLatitudeArray(sheet,i,j)[0];
                int lastCol =excelMergedRegionUtil.getMergedLatitudeArray(sheet,i,j)[1];
                int firstR =excelMergedRegionUtil.getMergedLatitudeArray(sheet,i,j)[2];
                int lastR =excelMergedRegionUtil.getMergedLatitudeArray(sheet,i,j)[3];
                // 获取合并单元格的二维坐标，同样也支持非合并单元格
                map.put("firstColumn",firstCol);
                map.put("lastColumn",lastCol);
                map.put("firstRow",firstR);
                map.put("lastRow",lastR);
                // 这里对起始行和结束行都加1是为了对下一行进行递归
                map.put("list",getCellBean(firstCol,lastCol,firstR+1,lastR+1));
                headList.add(map);
            }
        }
        return headList;
    }

    /**
     * 遍历List<Map<String,Object>>
     * @param cellBeanList
     */
    public void print(List<Map<String,Object>> cellBeanList){
        if (cellBeanList == null){
            return;
        }
        for (int i =0 ;i<cellBeanList.size();i++){
            System.out.println(cellBeanList.get(i).get("value"));
            print((List<Map<String,Object>>)cellBeanList.get(i).get("list"));
        }
    }

    public static void main(String[] args) {
        try {
           // File f = new File("F:\\航天信息(万万不要删)—日常薪资导入模板.xls");
            File f = new File("F:\\奖金筹划导入模板.xls");
            FileInputStream fis = new FileInputStream(f);
            BufferedInputStream bis = new BufferedInputStream(fis);
            POIFSFileSystem fs = new POIFSFileSystem(bis);
            HSSFWorkbook workbook = null;
            workbook = new HSSFWorkbook(fs);
            ExcelMergedRegionParser excelParserController = new ExcelMergedRegionParser(workbook,0,3);
          List<String> list = excelParserController.getHeaderList();
            for(int i = 0;i<list.size();i++){
                System.out.println(list.get(i)+", ");
            }
           /* System.out.println(excelParserController.getDataLineNum());
            excelParserController.print(excelParserController.getHeaderLayout());
            System.out.println(excelParserController.getDataLineNum());*/

        }catch (Exception e){
            e.printStackTrace();
        }


    }
}
