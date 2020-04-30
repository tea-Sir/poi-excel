package org.teasir.excel.poi;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.teasir.excel.poi.bean.Goods;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;


public class PoiWrite {

    //导出时设置表单，并生成表单
    private static XSSFSheet genSheet(XSSFWorkbook workbook, String sheetName) {
        //生成表单
        XSSFSheet sheet = workbook.createSheet(sheetName);
        //设置表单文本居中
        sheet.setHorizontallyCenter(true);
        sheet.setFitToPage(false);
        //打印时在底部右边显示文本页信息
        Footer footer = sheet.getFooter();
        footer.setRight("Page " + HeaderFooter.numPages() + " Of " + HeaderFooter.page());
        //打印时在头部右边显示Excel创建日期信息
        Header header = sheet.getHeader();
        header.setRight("Create Date " + HeaderFooter.date() + " " + HeaderFooter.time());
        //设置打印方式
        XSSFPrintSetup ps = sheet.getPrintSetup();
        ps.setLandscape(true); // true：横向打印，false：竖向打印 ，因为列数较多，推荐在打印时横向打印
        ps.setPaperSize(HSSFPrintSetup.A4_PAPERSIZE); //打印尺寸大小设置为A4纸大小
        return sheet;
    }

    //导出时创建文本样式
    private static XSSFCellStyle genContextStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);//文本水平居中显示
        style.setVerticalAlignment(VerticalAlignment.CENTER);//文本竖直居中显示
        style.setWrapText(true);//文本自动换行
        style.setBorderBottom(BorderStyle.THIN);//设置文本边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(new XSSFColor(java.awt.Color.BLACK));//设置文本边框颜色
        style.setBottomBorderColor(new XSSFColor(java.awt.Color.BLACK));
        style.setLeftBorderColor(new XSSFColor(java.awt.Color.BLACK));
        style.setRightBorderColor(new XSSFColor(java.awt.Color.BLACK));
        return style;
    }

    //导出时生成标题样式
    private static XSSFCellStyle genTitleStyle(XSSFWorkbook workbook) {

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        //标题居中，没有边框，所以这里没有设置边框，设置标题文字样式
        XSSFFont titleFont = workbook.createFont();
        titleFont.setBold(true);//加粗
        titleFont.setFontHeight((short) 10);//文字尺寸
        titleFont.setFontHeightInPoints((short) 10);
        style.setFont(titleFont);

        return style;
    }
/*
* 以XSSFWorkbook的格式导出.xlsx文件，适合于数据量不大的情况
* */
    public static ResponseEntity<byte[]> exportGoods2Excel(List<Goods> lgs) throws Exception {
        HttpHeaders headers;
        ByteArrayOutputStream baos;
       try {
            //1.创建Excel文档
            XSSFWorkbook workbook = new XSSFWorkbook();

            //创建Excel表单
            XSSFSheet sheet = genSheet(workbook, "处置信息表");
            //创建标题的显示样式
            XSSFCellStyle titleStyle = genTitleStyle(workbook);//创建标题样式
            titleStyle.setFillForegroundColor(IndexedColors.YELLOW.index);
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 冻结最左边的两列、冻结最上面的一行
            // 即：滚动横向滚动条时，左边的第一、二列固定不动;滚动纵向滚动条时，上面的第一行固定不动。
            sheet.createFreezePane(0, 1);
            // 创建第一行,作为header表头
            Row headerRow = sheet.createRow(0);
            //根据Excel列名长度，指定列名宽度
            for (int i = 0; i < 4; i++) {
                if (i == 1||i==3) {
                    sheet.setColumnWidth(i, 5000);
                } else {
                    sheet.setColumnWidth(i, 2500);
                }
            }

            //5.设置表头
            Cell cell0 = headerRow.createCell(0);
            cell0.setCellValue("商品编码");
            cell0.setCellStyle(titleStyle);
            Cell cell1 = headerRow.createCell(1);
            cell1.setCellValue("商品名字");
            cell1.setCellStyle(titleStyle);
            Cell cell2 = headerRow.createCell(2);
            cell2.setCellValue("商品数量");
            cell2.setCellStyle(titleStyle);
            Cell cell3 = headerRow.createCell(3);
            cell3.setCellValue("商品生产日期");
            cell3.setCellStyle(titleStyle);


            //6.装数据
            for (int i = 0; i < lgs.size(); i++) {
                Row row = sheet.createRow(i + 1);
                Goods goods = lgs.get(i);
                row.createCell(0).setCellValue(goods.getCode());
                row.createCell(1).setCellValue(goods.getName());
                row.createCell(2).setCellValue(goods.getAmount()==null?"":goods.getAmount().toString());
                row.createCell(3).setCellValue(goods.getDay());

            }
            headers = new HttpHeaders();
            headers.setContentDispositionFormData("attachment",
                    new String("商品信息表.xlsx".getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1));
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            baos = new ByteArrayOutputStream();
            workbook.write(baos);
        } catch (IOException e) {
            throw new Exception("导出异常！");
        }
        return new ResponseEntity<byte[]>(baos.toByteArray(), headers, HttpStatus.CREATED);
    }
    private static XSSFCellStyle getCellStyleHeader(SXSSFWorkbook sxssfWorkbook) {

        XSSFCellStyle style = (XSSFCellStyle) sxssfWorkbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        //标题居中，没有边框，所以这里没有设置边框，设置标题文字样式
        XSSFFont titleFont = (XSSFFont) sxssfWorkbook.createFont();
        titleFont.setBold(true);//加粗
        titleFont.setFontHeight((short) 10);//文字尺寸
        titleFont.setFontHeightInPoints((short) 10);
        style.setFont(titleFont);

        return style;

    }
    /*
     * 以SXSSFWorkbook的格式导出.xlsx文件，适合于数据量大的情况
     * */
    public static ResponseEntity<byte[]> exportGoods1Excel(List<Goods> lgs) throws Exception {
        HttpHeaders headers;
        ByteArrayOutputStream baos;
        SXSSFWorkbook workbook;
        try {
            //1.创建Excel文档
            workbook = new SXSSFWorkbook(5000);

            //创建Excel表单
            Sheet sheet = workbook.createSheet("商品信息表");
            // 冻结最左边的两列、冻结最上面的一行
            // 即：滚动横向滚动条时，左边的第一、二列固定不动;滚动纵向滚动条时，上面的第一行固定不动。
            sheet.createFreezePane(0, 1);

            // 创建第一行,作为header表头
            Row headerRow = sheet.createRow(0);

            //创建标题的显示样式
            XSSFCellStyle titleStyle = getCellStyleHeader(workbook);
            //创建标题样式
            titleStyle.setFillForegroundColor(IndexedColors.YELLOW.index);
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            //根据Excel列名长度，指定列名宽度
            for (int i = 0; i < 4; i++) {
                if (i == 1||i==3) {
                    sheet.setColumnWidth(i, 5000);
                } else {
                    sheet.setColumnWidth(i, 2500);
                }
            }

            //5.设置表头
            Cell cell0 = headerRow.createCell(0);
            cell0.setCellValue("商品编码");
            cell0.setCellStyle(titleStyle);
            Cell cell1 = headerRow.createCell(1);
            cell1.setCellValue("商品名字");
            cell1.setCellStyle(titleStyle);
            Cell cell2 = headerRow.createCell(2);
            cell2.setCellValue("商品数量");
            cell2.setCellStyle(titleStyle);
            Cell cell3 = headerRow.createCell(3);
            cell3.setCellValue("商品生产日期");
            cell3.setCellStyle(titleStyle);


            //6.装数据
            for (int i = 0; i < lgs.size(); i++) {
                Row row = sheet.createRow(i + 1);
                Goods goods = lgs.get(i);
                row.createCell(0).setCellValue(goods.getCode());
                row.createCell(1).setCellValue(goods.getName());
                row.createCell(2).setCellValue(goods.getAmount()==null?"":goods.getAmount().toString());
                row.createCell(3).setCellValue(goods.getDay());

            }
            headers = new HttpHeaders();
            headers.setContentDispositionFormData("attachment",
                    new String("商品信息表.xlsx".getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1));
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            baos = new ByteArrayOutputStream();
            workbook.write(baos);
        } catch (IOException e) {
            throw new Exception("导出异常！");
        }
        return new ResponseEntity<byte[]>(baos.toByteArray(), headers, HttpStatus.CREATED);
    }





}

