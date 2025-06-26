package personal.rushikesh.excelImporter.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import java.io.InputStream;
import java.util.*;

@Service
public class ExcelImportService {

    @Autowired
    private JdbcTemplate jdbcTemplate;

    public void importExcel(MultipartFile file) throws Exception {
        String tableName = getTableNameFromFile(file.getOriginalFilename());
        try (InputStream is = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new RuntimeException("No header row found");
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue().trim());
            }
            List<List<Object>> dataRows = new ArrayList<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                List<Object> rowData = new ArrayList<>();
                for (int j = 0; j < headers.size(); j++) {
                    rowData.add(getCellValue(row.getCell(j)));
                }
                dataRows.add(rowData);
            }
            Map<String, String> columnTypes = inferColumnTypes(sheet, headers);
            if (!tableExists(tableName)) {
                createTable(tableName, headers, columnTypes);
            }
            insertOrUpdateRows(tableName, headers, dataRows);
        }
    }

    private String getTableNameFromFile(String filename) {
        return filename == null ? "unknown_table" : filename.replaceAll("\\..*$", "");
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                if (d == (long) d) {
                    return (long) d; // Return as Long if no decimal part
                } else {
                    return d;
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            default:
                return null;
        }
    }

    private Map<String, String> inferColumnTypes(Sheet sheet, List<String> headers) {
        Map<String, String> types = new LinkedHashMap<>();
        types.put("Id", "BIGINT PRIMARY KEY AUTO_INCREMENT");
        for (int i = 1; i < headers.size(); i++) {
            String type = "VARCHAR(255)";
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if (row == null) continue;
                Cell cell = row.getCell(i);
                if (cell == null) continue;
                switch (cell.getCellType()) {
                    case NUMERIC:
                        type = "DOUBLE";
                        break;
                    case BOOLEAN:
                        type = "BOOLEAN";
                        break;
                    case STRING:
                        type = "VARCHAR(255)";
                        break;
                    default:
                        break;
                }
                if (!type.equals("VARCHAR(255)")) break;
            }
            types.put(headers.get(i), type);
        }
        return types;
    }

    private boolean tableExists(String tableName) {
        String sql = "SHOW TABLES LIKE ?";
        List<Map<String, Object>> result = jdbcTemplate.queryForList(sql, tableName);
        return !result.isEmpty();
    }

    private void createTable(String tableName, List<String> headers, Map<String, String> columnTypes) {
        StringBuilder sb = new StringBuilder("CREATE TABLE ").append(tableName).append(" (");
        for (String header : headers) {
            if (header.equalsIgnoreCase("Id")) {
                sb.append(header).append(" BIGINT PRIMARY KEY, ");
            } else {
                sb.append(header).append(" ").append(columnTypes.getOrDefault(header, "VARCHAR(255)")).append(", ");
            }
        }
        sb.setLength(sb.length() - 2);
        sb.append(")");
        jdbcTemplate.execute(sb.toString());
    }

    private void insertOrUpdateRows(String tableName, List<String> headers, List<List<Object>> dataRows) {
        String columns = String.join(", ", headers);
        String placeholders = String.join(", ", Collections.nCopies(headers.size(), "?"));

        // Build ON DUPLICATE KEY UPDATE clause (skip Id)
        StringBuilder updateClause = new StringBuilder();
        for (String header : headers) {
            if (!header.equalsIgnoreCase("Id")) {
                updateClause.append(header).append("=VALUES(").append(header).append("), ");
            }
        }
        if (updateClause.length() > 0) {
            updateClause.setLength(updateClause.length() - 2); // Remove last comma
        }

        String sql = "INSERT INTO " + tableName + " (" + columns + ") VALUES (" + placeholders + ")";
        if (updateClause.length() > 0) {
            sql += " ON DUPLICATE KEY UPDATE " + updateClause;
        }

        for (List<Object> row : dataRows) {
            jdbcTemplate.update(sql, row.toArray());
        }
    }
}