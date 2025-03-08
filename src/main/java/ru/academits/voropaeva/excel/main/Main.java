package ru.academits.voropaeva.excel.main;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import ru.academits.voropaeva.excel.Person;

import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        List<Person> persons = Arrays.asList(
                new Person("Лена", "Петрова", 19, "+79684664855"),
                new Person("Маша", "Иванова", 21, "+79554664859"),
                new Person("Александр", "Малышкин", 52, "+79564264875"),
                new Person("Евгения", "Ильина", 27, "+79574664852")
        );

        String[] headers = {"Имя", "Фамилия", "Возраст", "Телефон"};

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Список контактов");

        sheet.createFreezePane(0, 1);

        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        int i = 1;

        for (Person person : persons) {
            Row row = sheet.createRow(i);

            row.createCell(0).setCellValue(person.getName());
            row.createCell(1).setCellValue(person.getSurname());
            row.createCell(2).setCellValue(person.getAge());
            row.createCell(3).setCellValue(person.getTelephoneNumber());

            i++;
        }

        setCustomStyle(createBodyStyle(workbook), createHeadStyle(workbook), workbook);

        for (i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 500);
        }

        try (FileOutputStream fileOutputStream = new FileOutputStream("src/main/resources/persons.xlsx")) {
            workbook.write(fileOutputStream);
        } catch (Exception e) {
            System.out.println("Ошибка: " + e.getMessage());
        }

    }

    public static void setCustomStyle(CellStyle bodyStyle, CellStyle headStyle, Workbook workbook) {
        for (Sheet sheet : workbook) {
            int i = 0;

            for (Row row : sheet) {
                row.setHeight((short) 500);

                for (Cell cell : row) {
                    if (i == 0) {
                        cell.setCellStyle(headStyle);
                    } else {
                        cell.setCellStyle(bodyStyle);
                    }
                }

                ++i;
            }
        }
    }

    public static CellStyle createBodyStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);

        Font font = workbook.createFont();
        font.setFontName("News Gothic");
        font.setFontHeightInPoints((short) 13);
        style.setFont(font);

        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.LAVENDER.getIndex());

        return style;
    }

    public static CellStyle createHeadStyle(Workbook workbook) {
        CellStyle headStyle = workbook.createCellStyle();

        headStyle.setBorderBottom(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setBorderTop(BorderStyle.THIN);

        Font font = workbook.createFont();
        font.setFontName("News Gothic");
        font.setFontHeightInPoints((short) 13);
        font.setBold(true);
        headStyle.setFont(font);

        headStyle.setAlignment(HorizontalAlignment.LEFT);
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headStyle.setFillForegroundColor(IndexedColors.VIOLET.getIndex());

        return headStyle;
    }
}
