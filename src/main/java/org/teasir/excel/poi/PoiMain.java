package org.teasir.excel.poi;

import org.springframework.http.ResponseEntity;
import org.teasir.excel.poi.bean.Goods;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class PoiMain {
    public static void main(String[] args) throws Exception {
        String path1 = "E:\\\\StudyFileMap\\\\workspace\\\\poiExcel\\goods1.xlsx";
        /*
         * 1.HSSFWorkbook解析03版Excel和XSSFWorkboor解析07版Excel
         *返回集合中包含解析正确的数据，不符格式的记录，以及导入时间
         * */
        String operater = "admin";
        Map<String, List> map = PoiReader.importGoodsList(path1, operater);
        /*
        * 2.传入商品列表，返回字节数组输出流，前台导出Excel
        * */
        Goods goods=new Goods();
        goods.setCode("1");
        goods.setName("红烧牛肉面");
        goods.setAmount(new BigDecimal(120));
        goods.setDay("2020/4/30");
        List<Goods> goodsList=new ArrayList<Goods>();
        goodsList.add(goods);
        ResponseEntity<byte[]> responseEntity= PoiWrite.exportGoods1Excel(goodsList);
        ResponseEntity<byte[]> responseEntity2= PoiWrite.exportGoods2Excel(goodsList);
    }
}
