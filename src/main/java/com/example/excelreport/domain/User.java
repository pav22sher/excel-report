package com.example.excelreport.domain;

import com.example.excelreport.annotation.ReportAnnotation;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

@Data
@AllArgsConstructor
public class User {
    @ReportAnnotation(name = "ID")
    private Long id;
    @ReportAnnotation(name = "Полное имя")
    private String fullName;
    @ReportAnnotation(name = "ДР")
    private LocalDate birthday;
    @ReportAnnotation(name = "Сумма")
    private BigDecimal sum;
    @ReportAnnotation(name = "Создан")
    private LocalDateTime createdAt;

    public static List<User> getUsers() {
        return List.of(
                new User(0L, "Test0", LocalDate.now(), BigDecimal.ZERO, LocalDateTime.now()),
                new User(1L, "Test1", LocalDate.now(), BigDecimal.ONE, LocalDateTime.now()),
                new User(10L, "Test10", LocalDate.now(), BigDecimal.TEN, LocalDateTime.now())
        );
    }
}
