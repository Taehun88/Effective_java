package com.example.demo.dto;

import com.example.demo.poi.ExcelColumn;
import com.example.demo.poi.ExcelSheet;
import java.time.LocalDateTime;

@ExcelSheet(name = "Users")
public class UserExcelDto {

    @ExcelColumn(header = "ID", order = 1)
    private Long id;

    @ExcelColumn(header = "이름", order = 2)
    private String name;

    @ExcelColumn(header = "이메일", order = 3)
    private String email;

    @ExcelColumn(header = "가입일", order = 4)
    private LocalDateTime createdAt;
}
