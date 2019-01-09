import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public final class Util {

    private static SimpleDateFormat dateToString = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
    private static List<String> headerString = new ArrayList<>();

    public static List<ExcelModel> excelSorter(List<ExcelModel> excelModel) {

        Comparator<ExcelModel> cmp = Comparator.comparing(ExcelModel::getC1);
        cmp = cmp.thenComparing(ExcelModel::getC4);

        Stream<ExcelModel> excelModelStream = excelModel.stream().sorted(cmp);

        return excelModelStream.collect(Collectors.toList());
    }

    public static List<ExcelModel> excelGetter(File excelFile) {
        List<ExcelModel> excelModelList = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(excelFile)) {
            Sheet sheet = wb.getSheetAt(0);

            for (int i = 0; i < sheet.getLastRowNum(); i++) {
                Row r = sheet.getRow(i);
                ExcelModel excelModel = new ExcelModel();
                if (i == 0) {
                    for (int j = 0; j < 24; j++) {
                        if (!r.getCell(j).getCellType().equals(CellType.BLANK))
                            headerString.add(r.getCell(j).getStringCellValue());
                    }
                } else {
                    Cell c0 = r.getCell(1);
                    Cell c1 = r.getCell(4);
                    Cell c2 = r.getCell(20);

                    if (c0 != null && c1 != null && c2 != null && !c0.getCellType().equals(CellType.BLANK)) {
                        excelModel.setC0((int) r.getCell(0).getNumericCellValue());
                        excelModel.setC1(c0.getStringCellValue());
                        excelModel.setC2(r.getCell(2).getStringCellValue());
                        excelModel.setC3(r.getCell(3).getStringCellValue());
                        excelModel.setC4(c1.getNumericCellValue());
                        excelModel.setC5((int) r.getCell(5).getNumericCellValue());
                        excelModel.setC6(r.getCell(6).getStringCellValue());
                        excelModel.setC7(String.valueOf(r.getCell(7).getDateCellValue()));
                        excelModel.setC8(String.valueOf(r.getCell(8).getDateCellValue()));
                        excelModel.setC9(r.getCell(9).getStringCellValue());
                        excelModel.setC10(r.getCell(10).getStringCellValue());
                        excelModel.setC11(r.getCell(11).getStringCellValue());
                        excelModel.setC12((int) r.getCell(12).getNumericCellValue());
                        excelModel.setC13(r.getCell(13).getStringCellValue());
                        excelModel.setC14((int) r.getCell(14).getNumericCellValue());
                        excelModel.setC15(r.getCell(15).getStringCellValue());
                        excelModel.setC16(r.getCell(16).getNumericCellValue());
                        excelModel.setC17(r.getCell(17).getStringCellValue());
                        excelModel.setC18(r.getCell(18).getStringCellValue());
                        excelModel.setC19(r.getCell(19).getStringCellValue());
                        excelModel.setC20(c2.getNumericCellValue());
                        excelModel.setC21(r.getCell(21).getStringCellValue());
                        excelModel.setC22(r.getCell(22).getStringCellValue());
                        excelModel.setC23((int) r.getCell(23).getNumericCellValue());

                        excelModelList.add(excelModel);
                    }
                }
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return excelModelList;
    }

    public static void excelWriter(String excelName, List<ExcelModel> excelModelList) {
        try (Workbook workbook = new HSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sayfa1");
            CellStyle cellStyle = workbook.createCellStyle();
            CreationHelper createHelper = workbook.getCreationHelper();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m\\/d\\/yyyy"));
            Row rh = sheet.createRow(0);

            for (int i = 0; i < 24; i++) {
                rh.createCell(i).setCellValue(headerString.get(i));
            }

            int rowNum = 1;
            for (int j = 0; j < excelModelList.size(); j++) {
                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(excelModelList.get(j).getC0());
                row.createCell(1).setCellValue(excelModelList.get(j).getC1());
                row.createCell(2).setCellValue(excelModelList.get(j).getC2());
                row.createCell(3).setCellValue(excelModelList.get(j).getC3());
                row.createCell(4).setCellValue(excelModelList.get(j).getC4());
                row.createCell(5).setCellValue(excelModelList.get(j).getC5());
                row.createCell(6).setCellValue(excelModelList.get(j).getC6());
                row.createCell(7).setCellValue(dateConverter(excelModelList.get(j).getC7()));
                row.createCell(8).setCellValue(dateConverter(excelModelList.get(j).getC8()));
                row.createCell(9).setCellValue(excelModelList.get(j).getC9());
                row.createCell(10).setCellValue(excelModelList.get(j).getC10());
                row.createCell(11).setCellValue(excelModelList.get(j).getC11());
                row.createCell(12).setCellValue(excelModelList.get(j).getC12());
                row.createCell(13).setCellValue(excelModelList.get(j).getC13());
                row.createCell(14).setCellValue(excelModelList.get(j).getC14());
                row.createCell(15).setCellValue(excelModelList.get(j).getC15());
                row.createCell(16).setCellValue(excelModelList.get(j).getC16());
                row.createCell(17).setCellValue(excelModelList.get(j).getC17());
                row.createCell(18).setCellValue(excelModelList.get(j).getC18());
                row.createCell(19).setCellValue(excelModelList.get(j).getC19());
                row.createCell(20).setCellValue(excelModelList.get(j).getC20());
                row.createCell(21).setCellValue(excelModelList.get(j).getC21());
                row.createCell(22).setCellValue(excelModelList.get(j).getC22());
                row.createCell(23).setCellValue(excelModelList.get(j).getC23());
                row.getCell(7).setCellStyle(cellStyle);
                row.getCell(8).setCellStyle(cellStyle);

            }
            for (int i = 0; i < 24; i++) {
                sheet.autoSizeColumn(i);
            }
            try (FileOutputStream fo = new FileOutputStream(excelName + "_sorted.xls")) {
                workbook.write(fo);
            } catch (Exception ex) {
                System.out.println(ex.getMessage());
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    public static Date dateConverter(String s) {
        try {
            return dateToString.parse(s);
        } catch (ParseException parse) {
            System.out.println(parse.getMessage());
            return new Date();
        }
    }

}
