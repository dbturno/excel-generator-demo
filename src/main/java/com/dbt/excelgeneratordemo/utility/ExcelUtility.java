package com.dbt.excelgeneratordemo.utility;

import com.dbt.excelgeneratordemo.model.Book;
import com.dbt.excelgeneratordemo.model.Person;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

public class ExcelUtility {

    public static <T> void writeToExcel(String fileName,String sheetName, List<T> data) {
        OutputStream outputStream = null;
        XSSFWorkbook workbook = null;

        try {
            File file = new File(fileName);
            workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet(sheetName);

            List<String> fieldNames = getFieldNamesForClass(data.get(0).getClass());
            int rowCount = 0;
            int columnCount = 0;
            //Setup Header row
            Row row = sheet.createRow(rowCount++);
            for (String fieldName : fieldNames) {
                Cell cell = row.createCell(columnCount++);
                cell.setCellValue(fieldName);
            }

            //Setup Data row
            Class<? extends Object> aClass = data.get(0).getClass();
            for (T t: data) {
                row = sheet.createRow(rowCount++);
                columnCount = 0;
                for (String fieldName : fieldNames) {
                    Cell cell = row.createCell(columnCount);
                    Method method = null;
                    try {
                        method = aClass.getMethod("get" + capitalize(fieldName));
                    } catch (NoSuchMethodException nsme) {
                        method = aClass.getMethod("get" + fieldName);
                    }
                    Object value = method.invoke(t, (Object[]) null);
                    if (value != null) {
                        if (value instanceof String) {
                            cell.setCellValue((String) value);
                        } else if (value instanceof Long) {
                            cell.setCellValue((Long) value);
                        } else if (value instanceof Integer) {
                            cell.setCellValue((Integer) value);
                        } else if (value instanceof Double) {
                            cell.setCellValue((Double) value);
                        }
                    }
                    columnCount++;
                }
            }
            outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            outputStream.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private static String capitalize(String fieldName) {
        if (fieldName.length() == 0) {
            return fieldName;
        } else {
            return fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
        }
    }

    private static List<String> getFieldNamesForClass(Class<?> aClass) {
        List<String> fieldNames = new ArrayList<String>();
        Field[] fields = aClass.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            fieldNames.add(fields[i].getName());
        }
        return fieldNames;
    }

    //TODO: Design excel

    public static void main(String[] args) {
        List<Person> persons = new ArrayList<>();

        Person p1 = new Person("A", "a@roytuts.com", "Kolkata");
        Person p2 = new Person("B", "b@roytuts.com", "Mumbai");
        Person p3 = new Person("C", "c@roytuts.com", "Delhi");
        Person p4 = new Person("D", "d@roytuts.com", "Chennai");
        Person p5 = new Person("E", "e@roytuts.com", "Bangalore");
        Person p6 = new Person("F", "f@roytuts.com", "Hyderabad");

        persons.add(p1);
        persons.add(p2);
        persons.add(p3);
        persons.add(p4);
        persons.add(p5);
        persons.add(p6);

        writeToExcel("excel-person.xlsx", "Persons", persons);

        Book book1 = new Book("Head First Java", "Kathy Serria", 79);
        Book book2 = new Book("Effective Java", "Joshua Bloch", 36);
        Book book3 = new Book("Clean Code", "Robert Martin", 42);
        Book book4 = new Book("Thinking in Java", "Bruce Eckel", 35);

        List<Book> books = Arrays.asList(book1, book2, book3, book4);

        writeToExcel("excel-books.xlsx", "Books", books);
    }
}
