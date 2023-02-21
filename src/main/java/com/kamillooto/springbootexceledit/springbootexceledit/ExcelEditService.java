package com.kamillooto.springbootexceledit.springbootexceledit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

@Service
public class ExcelEditService {

    @Autowired
    private ProductRepository productRepository;

    public void updateExcelFile() {

        String excelFilePath = "C:\\Users\\medi\\Desktop\\kamil\\programowanie\\applications\\SpringbootExcelEdit\\spring-boot-excel-edit\\src\\main\\resources\\excelfile\\excel.xlsx";

        try {
            FileInputStream fileInputStream = new FileInputStream(excelFilePath);

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

            FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
