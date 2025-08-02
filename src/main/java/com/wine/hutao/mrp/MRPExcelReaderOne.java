package com.wine.hutao.mrp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @ClassName: MRPExcelReaderOne
 * @Description: 针对多个BOM的mrp计算需求值的汇总
 * 只需要输入BOM的id
 * @Author: binwine
 * @Date: 2025年8月02日 13:42
 */
public class MRPExcelReaderOne {
    public static void main(String[] args) {
//        String filePath = "C:\\Users\\binwine\\Desktop\\引出列表_MRP计算明细表_0802112552.xlsx";
        String filePath = "C:\\Users\\binwine\\Desktop\\引出列表_MRP计算明细表_0802115007.xlsx";
        String outputFilePath = "D:\\limit\\xx.xlsx";
        Workbook workbook;
        String mainBom1 = "1874599813908016128";
        String mainBom2 = "2224136567235038208";
        String mainBom3 = "2224135979613029376";
        String[] mainBomIDs = {mainBom1, mainBom2, mainBom3};
        Map<String, Double> bomIdMaterialMap = new HashMap<>();
        try {
            workbook = MRPExcelReader.readExcel(filePath);
            Sheet sheet = workbook.getSheetAt(0);
            int colNum = sheet.getLastRowNum();
//            System.out.println(colNum);
            // TODO: 实现您的逻辑
            //  获取索引21列BOMid，22列父项BOMid 跳过第0行标题
            // 第五列是 物料名称，9是需求数量 届是index
            for (int rowNum=0; rowNum < colNum; rowNum ++) {
                Row row = sheet.getRow(rowNum);
                if (rowNum ==0) continue;
//                System.out.println(row.getLastCellNum());
                String materialName = row.getCell(5).getStringCellValue();
                double materialNum = MRPExcelReader.getNumericCellValue(row.getCell(9));
                String parentBomId = "";
                String bomId = "";
                Cell cell35 = row.getCell(35);
                if (cell35!=null) {
                    bomId = row.getCell(35).getStringCellValue();
                }
                Cell cell36 = row.getCell(36);
                if (cell36!=null) {
                     parentBomId = row.getCell(36).getStringCellValue();
                }
                for (String mainBomID : mainBomIDs) {
                    if ((!bomId.equals("") && bomId.equals(mainBomID)) || parentBomId.equals(mainBomID)) {
                        // 一个BOM里面 但是物料名称和数量会多个 bomid++materialName为key
                        String key = mainBomID + "|" + materialName;
                        if (!bomIdMaterialMap.containsKey(key)) {
                            bomIdMaterialMap.put(key, materialNum);
                        } else {
                            double num = bomIdMaterialMap.get(key);
                            bomIdMaterialMap.put(key, num + materialNum);
                        }
                    }
                }
            }
            writeMapToExcel(bomIdMaterialMap, outputFilePath);
            System.out.println();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void writeMapToExcel(Map<String, Double> numberMap, String outputFilePath) throws IOException {
        // 1. 创建一个新的工作簿
        File file = new File(outputFilePath);
        Workbook workbook;
        if (file.exists()) {
//            System.out.println("📂 已存在，读取旧文件...");
            FileInputStream fis = new FileInputStream(file);
            workbook = new XSSFWorkbook(fis);
            fis.close();
        } else {
//            System.out.println("🆕 文件不存在，创建新文件...");
            workbook = new XSSFWorkbook();
        }
        int numberOfSheets = workbook.getNumberOfSheets() + 2;
        // 2. 创建一个工作表
        String sheetName = "汇总数据_" + numberOfSheets;
        Sheet sheet = workbook.createSheet(sheetName);
//        sheet = workbook.getSheetAt(numberOfSheets+1);
        // 3. 创建标题行
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("物料名称");
        headerRow.createCell(1).setCellValue("需求数量");

        // 4. 设置单元格样式（可选）
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        for (Cell cell : headerRow) {
            cell.setCellStyle(headerStyle);
        }

        // 5. 写入数据
        int rowNum = 1;
        for (Map.Entry<String, Double> entry : numberMap.entrySet()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(entry.getValue());
        }

        // 6. 自动调整列宽
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        // 7. 写入文件
        try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
            workbook.write(outputStream);
        }

        // 8. 关闭工作簿
        workbook.close();
    }

}
