import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        List<ExcelModel> unsortedExcelList = Util.excelGetter(new File("unsortedExcel.xlsx"));
        List<ExcelModel> sortedExcelList = Util.excelSorter(unsortedExcelList);
        Util.excelWriter(sortedExcelList);

    }
}
