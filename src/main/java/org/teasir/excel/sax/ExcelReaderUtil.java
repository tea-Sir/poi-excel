package org.teasir.excel.sax;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author teasir
 * @create 2020/4/30
 * @desc
 **/
public class ExcelReaderUtil {
    //excel2003扩展名
    public static final String EXCEL03_EXTENSION = ".xls";
    //excel2007扩展名
    public static final String EXCEL07_EXTENSION = ".xlsx";

    /**
     * @param sheetName
     * @param sheetIndex
     * @param curRow
     * @param cellList
     */
    /*

    /**
     * 读取数据,将每一条解析的记录存储在数组列表中
     */
    public static List<String[]> datas = new ArrayList<String[]>();

    public static void sendRows(String filePath, String sheetName, int sheetIndex, int curRow, List<String> cellList) {
        //每获取一条记录，即打印
        //在flume里每获取一条记录即发送，而不必缓存起来，可以大大减少内存的消耗，这里主要是针对flume读取大数据量excel来说的
        System.out.println(cellList);
        StringBuffer oneLineSb = new StringBuffer();
        oneLineSb.append(filePath);
        oneLineSb.append("--");
        oneLineSb.append("sheet" + sheetIndex);
        oneLineSb.append("::" + sheetName);//加上sheet名
        oneLineSb.append("--");
        oneLineSb.append("row" + curRow);
        oneLineSb.append("::");
        for (String cell : cellList) {
            oneLineSb.append(cell.trim());
            oneLineSb.append("|");
        }
        String oneLine = oneLineSb.toString();
        if (oneLine.endsWith("|")) {
            oneLine = oneLine.substring(0, oneLine.lastIndexOf("|"));
        }
        System.out.println(oneLine);
        //封装datas
        String[] cellArr;
        cellArr = cellList.toArray(new String[cellList.size()]);
        datas.add(cellArr.clone());
    }

    /*
     * 返回所有解析过的数据,也可以返回总行数
     * */
    public static List<String[]> readExcel(String fileName) throws Exception {
        int totalRows = 0;
        datas.clear();
        if (fileName.endsWith(".xls")) { //处理excel2003文件
            ExcelXlsReader excelXls = new ExcelXlsReader();
            totalRows = excelXls.process(fileName);
        } else if (fileName.endsWith(".xlsx")) {//处理excel2007文件
            ExcelXlsxReaderWithDefaultHandler excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler();
            totalRows = excelXlsxReader.process(fileName);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
        //打印总行数
        System.out.println("总行数：" + totalRows);
        return datas;
    }

    public static void copyToTemp(File file, String tmpDir) throws Exception {
        FileInputStream fis = new FileInputStream(file);
        File file1 = new File(tmpDir);
        if (file1.exists()) {
            file1.delete();
        }
        FileOutputStream fos = new FileOutputStream(tmpDir);
        byte[] b = new byte[1024];
        int n = 0;
        while ((n = fis.read(b)) != -1) {
            fos.write(b, 0, n);
        }
        fis.close();
        fos.close();
    }

    public static void main(String[] args) throws Exception {
        String path="E:\\\\StudyFileMap\\\\workspace\\\\poiExcel\\test4.xlsx";
        List<String[]> datas= ExcelReaderUtil.readExcel(path);

    }
    /*
    *如果是03版本的excel日期转化是正常的
    * 如果是07版本的excel日期转化后会转为mm/dd/yy格式
     * 目前解决方案是调用下列方法将mm/dd/yy转化为yyyy/mm/dd
     * */
    public String formatDate(String arrStr){
        String dateStr="";
        if(arrStr.matches("^([1-9]|0[1-9]|1[0-2])[/-]([1-9]|0[1-9]|[1-2][0-9]|3[0-1])[/-][0-9][0-9]$")){
            String[] ss = arrStr.split("\\p{Punct}");
            dateStr = ("20"+ss[2]) + "/"+ss[0]+ "/"+ss[1];
        }else {
            dateStr=arrStr;
        }
        return dateStr;
    }
}
