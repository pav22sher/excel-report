package com.example.excelreport.service;

import com.example.excelreport.annotation.ReportAnnotation;
import com.example.excelreport.domain.User;
import lombok.SneakyThrows;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Log4j2
@Service
public class ExportService {
    private static final DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm:ss");

    public XSSFWorkbook export1(List<User> users) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Users");
        int rowIndex = 0;
        XSSFRow row = sheet.createRow(rowIndex++);
        int cellIndex = 0;
        row.createCell(cellIndex).setCellValue("ID");
        row.createCell(++cellIndex).setCellValue("Полное имя");
        row.createCell(++cellIndex).setCellValue("ДР");
        row.createCell(++cellIndex).setCellValue("Сумма");
        row.createCell(++cellIndex).setCellValue("Создан");
        for (var user : users) {
            cellIndex = 0;
            row = sheet.createRow(rowIndex++);
            row.createCell(cellIndex).setCellValue(user.getId());
            row.createCell(++cellIndex).setCellValue(user.getFullName());
            row.createCell(++cellIndex).setCellValue(user.getBirthday().format(dateFormatter));
            row.createCell(++cellIndex).setCellValue(user.getSum().toString());
            row.createCell(++cellIndex).setCellValue(user.getCreatedAt().format(dateTimeFormatter));
        }
        for (int i = 0; i <= cellIndex; i++) {
            sheet.autoSizeColumn(i);
        }
        return workbook;
    }

    public XSSFWorkbook export2(List<?> rows) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        XSSFRow row;
        XSSFCell cell;
        int rowNum = 0;
        int cellNum = 0;
        if (!rows.isEmpty()) {
            Field[] fields = rows.get(0).getClass().getDeclaredFields();
            row = sheet.createRow(rowNum++);
            for (Field field : fields) {
                if (field.isAnnotationPresent(ReportAnnotation.class)) {
                    ReportAnnotation ra = field.getAnnotation(ReportAnnotation.class);
                    cell = row.createCell(cellNum++);
                    cell.setCellValue(ra.name());
                }
            }
            for (Object result : rows) {
                Class<?> aClass = result.getClass();
                fields = aClass.getDeclaredFields();
                cellNum = 0;
                row = sheet.createRow(rowNum++);
                for (Field field : fields) {
                    if (field.isAnnotationPresent(ReportAnnotation.class)) {
                        try {
                            cell = row.createCell(cellNum++);
                            cell.setCellValue(getCellValue(result, field));
                        }
                        catch (Exception e) {
                            log.error("Error while transforming object to xls report");
                        }
                    }
                }
            }
            for (int i = 0; i < cellNum; i++) {
                sheet.autoSizeColumn(i);
            }
        }
        return workbook;
    }

    @SneakyThrows
    private String getCellValue(Object result, Field field) {
        field.setAccessible(true);
        Object obj = field.get(result);
        if (obj instanceof LocalDate) {
            return ((LocalDate) obj).format(dateFormatter);
        }
        if (obj instanceof LocalDateTime) {
            return ((LocalDateTime) obj).format(dateTimeFormatter);
        }
        if (obj != null) {
            return obj.toString();
        } else {
            return "";
        }
    }
}
