import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        List<ExcelModel> unsortedExcelList = Util.excelGetter(new File("latest.xls"));
        List<ExcelModel> sortedExcelList = Util.excelSorter(unsortedExcelList);
        Util.excelWriter("latest", sortedExcelList);

    }
}
