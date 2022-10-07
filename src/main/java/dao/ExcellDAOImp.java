package dao;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcellDAOImp implements ExcelDAO{
    @Override
    public void createExcellInDisk(Workbook workbook, String path) throws IOException {
        try(FileOutputStream fileOutputStream = new FileOutputStream(path)){
            workbook.write(fileOutputStream);
            System.out.println("Fichero creado en " + path);
        }
    }
}
