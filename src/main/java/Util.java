
import com.setek.ExcelModel;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public final class Util {

    private static SimpleDateFormat dateToString = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
    private static SimpleDateFormat simpleDateFormat = new SimpleDateFormat("mm/dd/yyyy");
    private static List<String> headerString = new ArrayList<>();

    public static List<ExcelModel> excelSorter(List<ExcelModel> excelModel) {

        Comparator<ExcelModel> cmp = Comparator.comparing(ExcelModel::getC1);
        cmp = cmp.thenComparing(ExcelModel::getC4);

        Stream<ExcelModel> excelModelStream = excelModel.stream().sorted(cmp);

        return excelModelStream.collect(Collectors.toList());
    }

    public static List<ExcelModel> getExcelData(Sheet sheet) throws ParseException {
        int totalRows = sheet.getLastRowNum();
        int totalColumns = sheet.getRow(0).getLastCellNum();

        final DataFormatter dataFormatter = new DataFormatter();
        final Row headerRow = sheet.getRow(0);
        List<ExcelModel> excelModelList = new ArrayList<>();

        for (int i = 0; i < totalColumns; i++) {
            String headerValue = dataFormatter.formatCellValue(headerRow.getCell(i));
            if (headerValue.equals("")) {
                continue;
            }
            headerString.add(headerValue);
        }
        outer:
        for (int i = 1; i < totalRows; i++) {
            ExcelModel excelModel = new ExcelModel();
            Row row = sheet.getRow(i);
            for (int j = 0; j < totalColumns; j++) {
                String fieldValue = dataFormatter.formatCellValue(row.getCell(j));
                if (fieldValue.equals("")) {
                    continue outer;
                }
                switch (j) {
                    case 0: excelModel.setC0(fieldValue);break;
                    case 1: excelModel.setC1(fieldValue);break;
                    case 2: excelModel.setC2(fieldValue);break;
                    case 3: excelModel.setC3(fieldValue);break;
                    case 4: excelModel.setC4(Double.parseDouble(fieldValue));break;
                    case 5: excelModel.setC5(Integer.parseInt(fieldValue));break;
                    case 6: excelModel.setC6(fieldValue);break;
                    case 7:{
                        final Date date = simpleDateFormat.parse(fieldValue);
                        excelModel.setC7(simpleDateFormat.format(date));
                    break;}
                    case 8: {
                        final Date date = simpleDateFormat.parse(fieldValue);
                        excelModel.setC8(simpleDateFormat.format(date));break;}
                    case 9: excelModel.setC9(fieldValue);break;
                    case 10: excelModel.setC10(fieldValue);break;
                    case 11: excelModel.setC11(fieldValue);break;
                    case 12: excelModel.setC12(Integer.parseInt(fieldValue));break;
                    case 13: excelModel.setC13(fieldValue);break;
                    case 14: excelModel.setC14(Integer.parseInt(fieldValue));break;
                    case 15: excelModel.setC15(fieldValue);break;
                    case 16: excelModel.setC16(Double.parseDouble(fieldValue));break;
                    case 17: excelModel.setC17(fieldValue);break;
                    case 18: excelModel.setC18(fieldValue);break;
                    case 19: excelModel.setC19(fieldValue);break;
                    case 20: excelModel.setC20(Double.parseDouble(fieldValue));break;
                    case 21: excelModel.setC21(fieldValue);break;
                    case 22: excelModel.setC22(fieldValue);break;
                    case 23: excelModel.setC23(Integer.parseInt(fieldValue));break;
                    default:
                        break;
                }
            }
            excelModelList.add(excelModel);
        }
        return excelModelList;
    }

    public static List<ExcelModel> excelGetter(File excelFile) {
        try (Workbook wb = WorkbookFactory.create(excelFile)) {
            Sheet sheet = wb.getSheetAt(0);

            return getExcelData(sheet);

        } catch (IOException e) {
            System.out.println(e.getMessage());
            return new ArrayList<>();
        } catch (ParseException e) {
            e.printStackTrace();
            return new ArrayList<>();
        }
    }

    public static Path excelWriter(Path outputPath, String extension, List<ExcelModel> excelModelList) {
        final Path p = Paths.get(outputPath + "\\latest." + extension);

        try (Workbook workbook = new HSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sayfa1");
            CellStyle cellStyle = workbook.createCellStyle();
            CreationHelper createHelper = workbook.getCreationHelper();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m\\/d\\/yyyy"));
            Row rh = sheet.createRow(0);

            if (headerString.size() != 0)
                for (int i = 0; i < 24; i++) {
                    rh.createCell(i).setCellValue(headerString.get(i));
                }

            int rowNum = 1;
            for (ExcelModel anExcelModelList : excelModelList) {
                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(anExcelModelList.getC0());
                row.createCell(1).setCellValue(anExcelModelList.getC1());
                row.createCell(2).setCellValue(anExcelModelList.getC2());
                row.createCell(3).setCellValue(anExcelModelList.getC3());
                row.createCell(4).setCellValue(anExcelModelList.getC4());
                row.createCell(5).setCellValue(anExcelModelList.getC5());
                row.createCell(6).setCellValue(anExcelModelList.getC6());
                row.createCell(7).setCellValue(anExcelModelList.getC7());
                row.createCell(8).setCellValue(anExcelModelList.getC8());
                row.createCell(9).setCellValue(anExcelModelList.getC9());
                row.createCell(10).setCellValue(anExcelModelList.getC10());
                row.createCell(11).setCellValue(anExcelModelList.getC11());
                row.createCell(12).setCellValue(anExcelModelList.getC12());
                row.createCell(13).setCellValue(anExcelModelList.getC13());
                row.createCell(14).setCellValue(anExcelModelList.getC14());
                row.createCell(15).setCellValue(anExcelModelList.getC15());
                row.createCell(16).setCellValue(anExcelModelList.getC16());
                row.createCell(17).setCellValue(anExcelModelList.getC17());
                row.createCell(18).setCellValue(anExcelModelList.getC18());
                row.createCell(19).setCellValue(anExcelModelList.getC19());
                row.createCell(20).setCellValue(anExcelModelList.getC20());
                row.createCell(21).setCellValue(anExcelModelList.getC21());
                row.createCell(22).setCellValue(anExcelModelList.getC22());
                row.createCell(23).setCellValue(anExcelModelList.getC23());
                row.getCell(7).setCellStyle(cellStyle);
                row.getCell(8).setCellStyle(cellStyle);

            }
            for (int i = 0; i < 24; i++) {
                sheet.autoSizeColumn(i);
            }
            try (FileOutputStream fo = new FileOutputStream(p.toString())) {
                workbook.write(fo);
            } catch (Exception ex) {
                System.out.println(ex.getMessage());
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return p;
    }

    public static void excelMover(Path oldExcelDir, Path latestSortedExcel) throws IOException {
        final Path latestWillBeMoved = Util.newName(latestSortedExcel, Util.dateToFileName((latestSortedExcel.getFileName().toString().equals("latest.xls")) ? "xls" : "xlsx"));
        Files.move(latestWillBeMoved, oldExcelDir.resolve(latestWillBeMoved.getFileName()), REPLACE_EXISTING);
    }

    public static Date dateConverter(String s) {
        try {
            return dateToString.parse(s);
        } catch (ParseException parse) {
            System.out.println(parse.getMessage());
            return new Date();
        }
    }

    public static Path newName(Path oldName, String newNameString) throws IOException {
        return Files.move(oldName, oldName.resolveSibling(newNameString));
    }

    public static String dateToFileName(String fileType) {
        final Calendar calendar = Calendar.getInstance();
        final String pattern = "dd-MM-yyyy hh.mm.ss.SSS";
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);
        simpleDateFormat.setTimeZone(calendar.getTimeZone());
        return "latest_" + simpleDateFormat.format(calendar.getTime()) + "." + fileType;
    }

}