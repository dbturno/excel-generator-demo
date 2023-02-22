package com.dbt.excelgeneratordemo.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/v1/excel-generator/")
public class ExcelGeneratorController {

    @GetMapping("test")
    public String test() {
        return "Hello World!";
    }

}
