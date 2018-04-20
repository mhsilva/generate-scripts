package com.mhf.contacts.scripts;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReadingHelper {

    public List<ExcelPojo> createPojoFromExcel(String fileName) throws Exception {
        List<ExcelPojo> listPojo = new ArrayList<ExcelPojo>();
        Workbook workbook = new XSSFWorkbook(new FileInputStream(new File(fileName)));
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = sheet.iterator();
        iterator.next();
        while (iterator.hasNext()) {
            ExcelPojo excelPojo = new ExcelPojo();
            Row row = iterator.next();
            excelPojo.setSupplierType(getCellValue(row,0));
            excelPojo.setClassif(getCellValue(row,1));
            excelPojo.setSupplierCode(getCellValue(row,2));
            excelPojo.setSupplierName(getCellValue(row,3));
            excelPojo.setSapCountry(getCellValue(row, 4));
            excelPojo.setEmail(getCellValue(row, 6));
            excelPojo.setLoginId(getCellValue(row, 7).toLowerCase());
            excelPojo.setTelephone(getCellValue(row, 8));
            excelPojo.setIntensiveManagement(getCellValue(row, 9));
            excelPojo.setReceiveEmail(getCellValue(row, 10));
            List<String> companiesList = new ArrayList<String>();
            companiesList.add(getCellValue(row, 11));
            companiesList.add(getCellValue(row, 12));
            companiesList.add(getCellValue(row, 13));
            companiesList.add(getCellValue(row, 14));
            excelPojo.setCompanies(companiesList);
            listPojo.add(excelPojo);
        }
        return listPojo;
    }

    private String getCellValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return "";
        }
        cell.setCellType(CellType.STRING);
        String stringCellValue = cell.getStringCellValue();
        if ("Auma".equals(stringCellValue)) {
            stringCellValue = "Auma SA de CV";
        } else if ("PT".equals(stringCellValue)) {
            stringCellValue = "Plastic Tec";
        }
        return stringCellValue.replaceAll("\\n", "").replaceAll("\\r", "");
    }
}
