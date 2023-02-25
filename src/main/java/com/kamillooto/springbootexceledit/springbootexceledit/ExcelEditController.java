package com.kamillooto.springbootexceledit.springbootexceledit;

import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

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

    @GetMapping(path = "/download_updated_excel")
    public void editAndDownloadExcelReport(HttpServletResponse response) {
        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss");
        LocalDateTime now = LocalDateTime.now();
        String name = dtf.format(now);
        String headerValue = "attachment;filename=new-" + name + ".xlsx";
        response.setHeader(headerKey, headerValue);
        excelEditService.editAndDownloadExcelReport(response);
    }
}
