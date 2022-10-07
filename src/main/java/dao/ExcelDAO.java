package dao;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;

public interface ExcelDAO {

    void createExcellInDisk(Workbook workbook, String path) throws IOException;
}
