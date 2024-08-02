package ru.academits.voropaeva.main;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import ru.academits.voropaeva.Person;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        List<Person> persons = new ArrayList<>();

        persons.add(new Person("Лена", "Петрова", 19, "+79684664855"));
        persons.add(new Person("Маша", "Иванова", 21, "+79554664859"));
        persons.add(new Person("Александр", "Малышкин", 52, "+79564264875"));
        persons.add(new Person("Евгения", "Ильина", 27, "+79574664852"));

        String[] headers = {"Имя", "Фамилия", "Возраст", "Телефон"};

        File personsFile = new File("src/main/resources/persons.xls");

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Список контактов");

        sheet.createFreezePane(0, 1);

        Row rowHead = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++) {
            rowHead.createCell(i).setCellValue(headers[i]);
        }

        int indexRow = 1;

        for (Person person : persons) {
            Row row = sheet.createRow(indexRow);
            indexRow++;

            row.createCell(0).setCellValue(person.getName());
            row.createCell(1).setCellValue(person.getSurname());
            row.createCell(2).setCellValue(person.getAge());
            row.createCell(3).setCellValue(person.getTelephoneNumber());
        }

        setCustomStyle(createCustomStyleBody(workbook), createCustomStyleHead(workbook), workbook);

        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, (sheet.getColumnWidth(i) + 500));
        }

        FileOutputStream fileOutputStream = new FileOutputStream(personsFile);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public static void setCustomStyle(CellStyle bodyStyle, CellStyle headStyle, Workbook workbook) {
        for (Sheet sheet : workbook) {
            int indexRow = 0;

            for (Row row : sheet) {
                row.setHeight((short) 500);

                for (Cell cell : row) {
                    cell.setCellStyle(bodyStyle);

                    if (indexRow == 0) {
                        cell.setCellStyle(headStyle);
                    }
                }

                ++indexRow;
            }
        }
    }

    public static CellStyle createCustomStyleBody(Workbook workbook) {
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

    public static CellStyle createCustomStyleHead(Workbook workbook) {
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
