package com.kamillooto.springbootexceledit.springbootexceledit;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static org.apache.poi.ss.util.CellUtil.getRow;

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

    public void generatePdfFromXlsFile() throws Exception {
        String excelFilePath = "C:\\Users\\medi\\Desktop\\kamil\\programowanie\\applications\\SpringbootExcelEdit\\spring-boot-excel-edit\\src\\main\\resources\\excelfile\\excel1.xlsx";
        String pdfDirectoryToSave = "C:\\Users\\medi\\Desktop\\kamil\\programowanie\\applications\\SpringbootExcelEdit\\spring-boot-excel-edit\\src\\main\\resources\\excelfile\\";

        com.grapecity.documents.excel.Workbook workbook = new com.grapecity.documents.excel.Workbook();
        workbook.open(excelFilePath);

        workbook.save(pdfDirectoryToSave + "pdf1.pdf");

    }
}
