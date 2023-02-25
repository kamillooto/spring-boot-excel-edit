package com.kamillooto.springbootexceledit.springbootexceledit;

import com.grapecity.datavisualization.chart.typescript.Date;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ExcelEditService {

    @Autowired
    private ProductRepository productRepository;

    public void updateExcelFile() {

        String excelTemplateFilePath = "src\\main\\resources\\excelfile\\excel.xlsx";
        String excelFileToSavePath = "src\\main\\resources\\excelfile\\";
        try {
            FileInputStream fileInputStream = new FileInputStream(excelTemplateFilePath);

            Workbook workbook = WorkbookFactory.create(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);

            int firstRowNumberToUpdateToFile = 5;

            List<Product> products = productRepository.findAll();

            for (Product product : products) {
                Row dataRow = sheet.getRow(firstRowNumberToUpdateToFile++);
                dataRow.getCell(0).setCellValue(product.getId());
                dataRow.getCell(1).setCellValue(product.getProductName());
                dataRow.getCell(2).setCellValue(product.getPrice());
            }

            fileInputStream.close();

            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss");
            LocalDateTime now = LocalDateTime.now();
            String date = dtf.format(now);
            System.out.println(date);
            FileOutputStream fileOutputStream = new FileOutputStream(excelFileToSavePath + date + ".xlsx");
            workbook.write(fileOutputStream);
            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void generatePdfFromXlsFile() throws Exception {
        String excelFilePath = "C:\\Users\\medi\\Desktop\\kamil\\programowanie\\applications\\SpringbootExcelEdit\\spring-boot-excel-edit\\src\\main\\resources\\excelfile\\excel1.xlsx";
        String pdfDirectoryToSave = "C:\\Users\\medi\\Desktop\\kamil\\programowanie\\applications\\SpringbootExcelEdit\\spring-boot-excel-edit\\src\\main\\resources\\excelfile\\";

        com.grapecity.documents.excel.Workbook workbook = new com.grapecity.documents.excel.Workbook();
        workbook.open(excelFilePath);

        workbook.save(pdfDirectoryToSave + "pdf1.pdf");

    }
}
