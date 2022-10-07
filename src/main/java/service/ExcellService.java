package service;

import dao.ExcelDAO;
import dao.ExcellDAOImp;
import entity.ComicInfoEntity;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
public class ExcellService {

    ExcelDAO excelDAO = new ExcellDAOImp();

    public  void creaExcellSuperHeroes(String filePath){
        List<ComicInfoEntity> comicInfoEntitiesList = loadInfo();
        try (Workbook workbook = new XSSFWorkbook()){
            Sheet sheet = workbook.createSheet();
            createHeader(sheet,workbook.createCellStyle(),workbook.createFont());
            createRow(comicInfoEntitiesList, sheet);
            excelDAO.createExcellInDisk(workbook, filePath);
        } catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    private void createRow(List<ComicInfoEntity> comicInfoEntitiesList, Sheet sheet) {
        for (int i = 0; i < comicInfoEntitiesList.size() ; i++) {
            ComicInfoEntity comicInfoEntity = comicInfoEntitiesList.get(i);
            Row row = sheet.createRow(i + 1);
            Cell cell = row.createCell(0);
            RichTextString text = new XSSFRichTextString(comicInfoEntity.getSuperHeroe());
            cell.setCellValue(text);

            cell = row.createCell(1);
            text = new XSSFRichTextString(comicInfoEntity.getCompany());
            cell.setCellValue(text);

            cell = row.createCell(2);
            text = new XSSFRichTextString(comicInfoEntity.getCreator());
            cell.setCellValue(text);

            cell = row.createCell(3);
            text = new XSSFRichTextString(comicInfoEntity.getFirstApparition());
            cell.setCellValue(text);

             cell = row.createCell(4);
             text = new XSSFRichTextString(comicInfoEntity.getDateApparition());
             cell.setCellValue(text);
        }
    }

    private void createHeader (Sheet sheet, CellStyle cellStyle, Font font) {
        cellStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setBold(true);

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        RichTextString text = new XSSFRichTextString("Super Heroe");
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);

        cell = row.createCell(1);
        text = new XSSFRichTextString("Compañía");
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);

        cell = row.createCell(2);
        text = new XSSFRichTextString("Creador");
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);

        cell = row.createCell(3);
        text = new XSSFRichTextString("Primera Aparición");
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);

        cell = row.createCell(4);
        text = new XSSFRichTextString("Fecha Primera Aparición");
        cell.setCellValue(text);
        cell.setCellStyle(cellStyle);
    }

    private List<ComicInfoEntity> loadInfo() {

        List<ComicInfoEntity> comicInfoEntitiesList = new ArrayList<>();
        comicInfoEntitiesList.add(new ComicInfoEntity("Spiderman", "Marvel", "Stan Lee y Steve Ditko", "Amazing Fantasy #15", "Agosto de 1962" ));
        comicInfoEntitiesList.add(new ComicInfoEntity("Superman", "DC", "Jerry Siegel y Joe Shuster", "Action Comics #1", "Junio de 1938"));
        comicInfoEntitiesList.add(new ComicInfoEntity("Batman", "DC", "Bob Kane y Bill Finger", "Detective Comics #27", "Marzo de 1939"));
        comicInfoEntitiesList.add(new ComicInfoEntity("Iron Man", "Marvel", "Stan Lee, Larry Lieber, Don Heck y Jack Kirby", "Tales of Suspense #39", "Marzo de 1963"));

        return comicInfoEntitiesList;
    }
}
