package personal.rushikesh.excelImporter.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import personal.rushikesh.excelImporter.entity.Employee;

public interface EmployeeRepository extends JpaRepository<Employee, Long> {
}
