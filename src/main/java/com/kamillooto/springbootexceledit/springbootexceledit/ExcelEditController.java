package com.kamillooto.springbootexceledit.springbootexceledit;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
public class ExcelEditController {

    @Autowired
    private ExcelEditService excelEditService;

    @GetMapping(path = "/update_excel")
    public void editExcelReport() {
        excelEditService.updateExcelFile();
    }

    @GetMapping(path = "generate_pdf")
    public void generatePdf() throws Exception {
        excelEditService.generatePdfFromXlsFile();
    }
}
