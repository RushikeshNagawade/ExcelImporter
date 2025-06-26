package personal.rushikesh.excelImporter.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import personal.rushikesh.excelImporter.service.EmployeeService;

@RestController
@RequestMapping("/api/employees")
public class EmployeeController {
    @Autowired
    private EmployeeService service;

    @PostMapping("/upload")
    public ResponseEntity<String> uploadExcel(@RequestParam("file") MultipartFile file) {
        try {
            service.saveEmployeesFromExcel(file);
            return ResponseEntity.ok("Employees uploaded successfully");
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("Upload failed: " + e.getMessage());
        }
    }
}
