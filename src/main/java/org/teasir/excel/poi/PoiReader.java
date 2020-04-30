package org.teasir.excel.poi;



import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.multipart.MultipartFile;
import org.teasir.excel.poi.bean.Goods;
import org.teasir.excel.poi.bean.GoodsImport;
import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/*
* 适用于数据量少于3000条的excel导入
* */
public class PoiReader {


    private final static String xls = "xls";
    private final static String xlsx = "xlsx";


    //导入时检查文件
    private static void checkFile(MultipartFile file) throws Exception {
        //判断文件是否存在
        if (null == file) {
            throw new Exception("文件不存在！");
        }
        //获得文件名
        String fileName = file.getOriginalFilename();
        //判断文件是否是excel文件
        if (!fileName.endsWith(xls) && !fileName.endsWith(xlsx)) {
            throw new Exception("不是excel文件！");
        }
    }

    //导入时获得工作簿
    private static Workbook getWorkBook(String  path) throws Exception {
        //获得文件名
        String fileName = path;
        //创建Workbook工作薄对象，表示整个excel
        Workbook workbook = null;
        try {
            //获取excel文件的io流
            InputStream is = new FileInputStream(path);
            //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
            if (fileName.endsWith(xls)) {
                //2003
                workbook = new HSSFWorkbook(is);
            } else if (fileName.endsWith(xlsx)) {
                //2007
                workbook = new XSSFWorkbook(is);
            }
        } catch (IOException e) {
            throw new Exception("创建工作簿失败！");
        }
        return workbook;
    }
    //处理单元格的null并获得单元格的值
    private static String getCellValue(Cell cell) {
        //判断是否为null或空串
        if (cell == null || "".equals(cell.toString().trim())) {
            return "";
        }
        String cellValue = "";
        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            //判断日期类型
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                Date dateCellValue = cell.getDateCellValue();
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
                cellValue = sdf.format(dateCellValue);
            } else {
                cellValue = new DecimalFormat("#.####").format(cell.getNumericCellValue());

            }

        } else if (cell.getCellTypeEnum() == CellType.STRING) {
            cellValue = String.valueOf(cell.getStringCellValue());
        } else if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellTypeEnum() == CellType.ERROR) {
            cellValue = "错误类型";
        } else {
            cellValue = "";
        }
        return cellValue;
    }



    //判断该行是不是空行
    public static boolean isRowEmpty(Row row) {
        Boolean flag = false;
        for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellTypeEnum() != CellType.BLANK) {
                flag = true;
                break;
            }
        }
        return flag;
    }



    public static Map<String, List> importGoodsList(String  path, String operster) throws Exception {
        //checkFile(file);
        Workbook workbook = getWorkBook(path);
        List<Goods> goodsList = new ArrayList<Goods>();
        List<GoodsImport> imports = new ArrayList<GoodsImport>();
        List<String> importDate = new ArrayList<String>();
        Map<String, List> map = new HashMap<String, List>(3);
        int numberOfSheets = workbook.getNumberOfSheets();
        //导入校验
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd HH:mm:ss");
        String importTime = sdf.format(new Date());
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();

            Goods goods;
            GoodsImport gi;
            for (int j = 0; j < physicalNumberOfRows; j++) {

                if (j == 0) {
                    continue;//标题行
                }
                Row row = sheet.getRow(j);
                if (row == null) {
                    continue;//没数据
                }
                int physicalNumberOfCells = row.getLastCellNum();
                goods = new Goods();
                //导入校验
                Boolean flag = true;
                gi = new GoodsImport();
                gi.setRowNum(j);
                gi.setOperator(operster);
                gi.setImportDate(importTime);
                StringBuffer sbf = new StringBuffer("");
                if (isRowEmpty(row)) {
                    for (int k = 0; k < physicalNumberOfCells; k++) {
                        Cell cell = row.getCell(k);
                        String cellValue = getCellValue(cell);

                        if (cellValue == null) {
                            cellValue = "";
                        }
                        switch (k) {
                            case 0:
                                if ("".equals(cellValue)) {
                                    sbf.append("-商品代码:不为空");
                                    continue;
                                } else if (!cellValue.matches("^[A-Za-z0-9_]{0,8}$")) {
                                    sbf.append("-商品代码:应为数字、字母、下划线组合并且长度不超过8位");
                                    continue;
                                }
                                goods.setCode(cellValue);
                                break;
                            case 1:
                                if ("".equals(cellValue)) {
                                    sbf.append("-商品名字:不为空");
                                    continue;
                                } else if (cellValue.length() > 50) {
                                    sbf.append("-商品名字:长度不超过50位");
                                    continue;
                                }
                                goods.setName(cellValue);
                                break;


                            /*BigDecimal数值类型*/
                            case 2:
                                if (!"".equals(cellValue)) {
                                    if (!cellValue.matches("^(([1-9]{1}\\d{0,7})|(0{1}))(\\.\\d{0,2})?$")) {
                                        sbf.append("-商品数量应为数字类型并且长度最大为10位，保留两位小数");
                                        continue;
                                    }
                                    goods.setAmount(new BigDecimal(cellValue));
                                }
                                break;
                                /*日期格式*/
                            case 3:
                                if (!"".equals(cellValue)) {
                                    if (!cellValue.matches("^[1-9]\\d{3}[/-](0[1-9]|1[0-2])[/-](0[1-9]|[1-2][0-9]|3[0-1])$")) {
                                        sbf.append("-商品生产日期:应为日期格式");
                                        continue;
                                    } else {
                                        goods.setDay(cellValue.substring(0, 4) + cellValue.substring(5, 7) + cellValue.substring(8, 10));
                                    }
                                } else {
                                    goods.setDay(cellValue);
                                }
                                break;
                        }

                    }
                    if (!"".equals(sbf.toString())) {
                        flag = false;
                    }

                    if (flag) {
                        goodsList.add(goods);
                    } else {
                        gi.setCode(goods.getCode());
                        gi.setName(goods.getName());
                        if (sbf.toString().length() > 600) {
                            gi.setErrReason(sbf.substring(0, 600));
                        } else {
                            gi.setErrReason(sbf.toString());
                        }
                        imports.add(gi);
                    }
                }
            }
        }
        importDate.add(importTime);
        map.put("importDate", importDate);
        map.put("import", imports);
        map.put("goods", goodsList);
        return map;
    }



}
