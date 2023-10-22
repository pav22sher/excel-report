package com.example.excelreport.controller;

import com.example.excelreport.domain.User;
import com.example.excelreport.service.ExportService;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import lombok.SneakyThrows;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.util.List;

@Controller
public class ExportController {
    @Autowired
    private ExportService exportService;

    @GetMapping("/export1")
    @SneakyThrows
    public void export1(HttpServletResponse response) {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=users1.xlsx");
        List<User> users = User.getUsers();
        XSSFWorkbook workbook = exportService.export1(users);
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    @GetMapping("/export2")
    @SneakyThrows
    public void export2(HttpServletResponse response) {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=users2.xlsx");
        List<User> users = User.getUsers();
        XSSFWorkbook workbook = exportService.export2(users);
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }
}
