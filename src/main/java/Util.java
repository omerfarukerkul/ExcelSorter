import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public final class Util {

    public static List<ExcelModel> excelSorter(List<ExcelModel> excelModel) {

        Comparator<ExcelModel> cmp = Comparator.comparing(ExcelModel::getC0);
        cmp = cmp.thenComparing(ExcelModel::getC1);

        Stream<ExcelModel> excelModelStream = excelModel.stream().sorted(cmp);

        return excelModelStream.collect(Collectors.toList());
    }

    public static List<ExcelModel> excelGetter(File excelFile) {
        List<ExcelModel> excelModelList = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(excelFile)) {
            Sheet sheet = wb.getSheetAt(0);

            for (Row r : sheet) {
                ExcelModel excelModel = new ExcelModel();

                Cell c0 = r.getCell(0);
                Cell c1 = r.getCell(1);
                Cell c2 = r.getCell(2);

                if (c0 != null && c1 != null && c2 != null) {
                    if (c0.getCellType() == CellType.STRING && c1.getCellType() == CellType.STRING && c2.getCellType() == CellType.NUMERIC) {
                        excelModel.setC0(c0.getStringCellValue());
                        excelModel.setC1(c1.getStringCellValue());
                        excelModel.setC2((int) c2.getNumericCellValue());
                        excelModelList.add(excelModel);
                    }
                }
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return excelModelList;
    }

    public static void excelWriter(List<ExcelModel> excelModelList) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sayfa1");

            int rowNum = 0;
            for (ExcelModel excelModel : excelModelList) {
                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(excelModel.getC0());
                row.createCell(1).setCellValue(excelModel.getC1());
                row.createCell(2).setCellValue(excelModel.getC2());
            }

            try (FileOutputStream fo = new FileOutputStream("sortedExcel.xlsx")) {
                workbook.write(fo);
            } catch (Exception ex) {
                System.out.println(ex.getMessage());
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

}
