package com.wine.hutao.mrp;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: MRPExcelReader
 * @Description:
 * @Author: binwine
 * @Date: 2025年7月31日 19:21
 */


public class MRPExcelReader {
    public static void main(String[] args) {
        // 批次
        int piCi = 40;
        // 批数量
        int pNumber = 72000;
        double computedResult = 720000;
        double computedResult1 = 15*48000;
        String filePath = "D:\\limit\\example.xlsx";

        try {
            // 1. 读取Excel文件
            Workbook workbook = readExcel(filePath);
            Sheet sheet = workbook.getSheetAt(4);

            CellStyle referenceStyle = sheet.getRow(0).getCell(4).getCellStyle();
            // 创建相同样式的新样式对象
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(referenceStyle);

            // 创建黄色填充样式
            CellStyle yellowStyle = workbook.createCellStyle();
            yellowStyle.cloneStyleFrom(referenceStyle);
            yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // 边框设置（四边）
            yellowStyle.setBorderTop(BorderStyle.THIN);
            yellowStyle.setBorderBottom(BorderStyle.THIN);
            yellowStyle.setBorderLeft(BorderStyle.THIN);
            yellowStyle.setBorderRight(BorderStyle.THIN);
            yellowStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            yellowStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            yellowStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            yellowStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            // 2. 存储读取的数据和计算结果
//            List<Double> column8Data = new ArrayList<>(); // H列
//            List<Double> column9Data = new ArrayList<>(); // I列
            List<Double> calculatedResults = new ArrayList<>(); // 计算结果
            Map<String, Double> zuJianMap = new HashMap<>();
            // 3. 遍历行并处理数据
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    row = sheet.createRow(rowNum);
                }
                // 读取E列数据
                // 获取参考单元格的样式
                // 获取下一行E列的值
                if (row.getCell(4)==null || row.getCell(4).getStringCellValue().equals("")) {
                    // 获取下一行B列的值 将map里面的值给computedResult
                    if (rowNum ==sheet.getLastRowNum()) break;
                    Row row1 = sheet.getRow(rowNum + 1);
                    String key = row1.getCell(1).getStringCellValue();
                    if (zuJianMap.containsKey(key)) {
                        computedResult = zuJianMap.get(key);
                    }
                    continue;
                }
                Cell referenceCell = row.getCell(4);
                String componentV1 = referenceCell.getStringCellValue();
                // 读取H列和I列数据
                double fenZi = getNumericCellValue(row.getCell(7)); // H列
                double fenMu = getNumericCellValue(row.getCell(8)); // I列
//                column8Data.add(fenZi);
//                column9Data.add(fenMu);
                // 计算MRP结果（替换为您的实际计算逻辑）
                String stringCellValue = row.getCell(0).getStringCellValue();
                if (stringCellValue.equals("BOM-A01103000045")) {
                    computedResult = computedResult1;
                }
                double calculatedValue = calculateMrp(fenZi, fenMu, computedResult);
                calculatedResults.add(calculatedValue);
                if (!zuJianMap.containsKey(componentV1)) {
                    zuJianMap.put(componentV1, calculatedValue);
                } else {
                    zuJianMap.put(componentV1, zuJianMap.get(componentV1) + calculatedValue);
                }
                // 将结果写入J列（第10列，索引9）
                Cell resultCell = row.createCell(9); // J列
                resultCell.setCellValue(calculatedValue);
                resultCell.setCellStyle(newStyle);
                if (row.getCell(6).getBooleanCellValue()) {
                    resultCell.setCellStyle(yellowStyle);
                }
            }

            // 4. 保存修改到原文件
            saveWorkbook(workbook, filePath);
            System.out.println("MRP计算结果已成功写入原文件的J列");

            // 5. 关闭工作簿
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 读取Excel文件
    public static Workbook readExcel(String filePath) throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(new File(filePath))) {
            return WorkbookFactory.create(fileInputStream);
        }
    }

    // 保存工作簿到原文件
    public static void saveWorkbook(Workbook workbook, String filePath) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }
    }

    // 安全获取数值型单元格值
    public static double getNumericCellValue(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return 0.0;
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue());
            } catch (NumberFormatException e) {
                return 0.0;
            }
        }
        return 0.0;
    }

    // MRP计算逻辑（示例）
    public static double calculateMrp(double zi, double mu, double computedResult) {
        return computedResult / mu * zi;
    }
}
