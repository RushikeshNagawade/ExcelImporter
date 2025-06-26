package personal.rushikesh.excelImporter.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import personal.rushikesh.excelImporter.entity.Employee;
import personal.rushikesh.excelImporter.repository.EmployeeRepository;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

@Service
public class EmployeeService {
    @Autowired
    private EmployeeRepository repository;

    public void saveEmployeesFromExcel(MultipartFile file) throws Exception {
        try (InputStream is = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<String, Integer> headerMap = new HashMap<>();
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                headerMap.put(cell.getStringCellValue().trim().toLowerCase(), cell.getColumnIndex());
            }
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Employee emp = new Employee();
                if (headerMap.containsKey("name")) {
                    emp.setName(getCellValue(row.getCell(headerMap.get("name"))));
                }
                if (headerMap.containsKey("email")) {
                    emp.setEmail(getCellValue(row.getCell(headerMap.get("email"))));
                }
                if (headerMap.containsKey("department")) {
                    emp.setDepartment(getCellValue(row.getCell(headerMap.get("department"))));
                }
                repository.save(emp);
            }
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue();
        if (cell.getCellType() == CellType.NUMERIC) return String.valueOf(cell.getNumericCellValue());
        if (cell.getCellType() == CellType.BOOLEAN) return String.valueOf(cell.getBooleanCellValue());
        return null;
    }
}