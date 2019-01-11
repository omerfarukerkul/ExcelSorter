import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.List;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public class Main {

    public static void main(String[] args) throws IOException, InterruptedException {
        final Path excelDir = Paths.get("C:", "Users", "OFE", "Desktop", "excelPath");
        final Path oldExcelDir = excelDir.getParent().resolve("oldExcelPath");
        List<Path> firstPathList = new ArrayList<>();
        List<Path> secondPathList = new ArrayList<>();

        if (!Files.exists(excelDir)) {
            Files.createDirectories(excelDir);
            System.out.println("Excel directory has been created.");
        }
        if (!Files.exists(oldExcelDir)) {
            Files.createDirectories(oldExcelDir);
            System.out.println("Sorted excel directory has been created.");
        }

        Files.walkFileTree(excelDir, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                    firstPathList.add(file);

                return FileVisitResult.CONTINUE;
            }
        });

        for (Path path : firstPathList) {
            if (path.getFileName().toString().equals("latest.xls") || path.getFileName().toString().equals("latest.xlsx")) {
                // TODO: 11.01.2019  Dosya burada işlenecek.
                final Path newName = Util.newName(path, Util.dateToFileName((path.getFileName().toString().equals("latest.xls")) ? "xls" : "xlsx"));
                Files.move(newName, oldExcelDir.resolve(newName.getFileName()), REPLACE_EXISTING);

            }
        }

        Files.walkFileTree(excelDir, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                secondPathList.add(file);

                return FileVisitResult.CONTINUE;
            }
        });

        for (Path path : secondPathList) {
            final List<ExcelModel> excelModelList = Util.excelGetter(new File(path.toString()));
            final List<ExcelModel> sortedExcelList = Util.excelSorter(excelModelList);
            final Path latestSortedExcel = Util.excelWriter(excelDir,(path.getFileName().toString().contains("xlsx")) ? "xlsx" : "xls", sortedExcelList);
            Files.delete(path);
            Thread.sleep(1000);
            // TODO: 11.01.2019  Dosya burada işlenecek.
            final Path latestWillBeMoved = Util.newName(latestSortedExcel, Util.dateToFileName((latestSortedExcel.getFileName().toString().equals("latest.xls")) ? "xls" : "xlsx"));
            Files.move(latestWillBeMoved, oldExcelDir.resolve(latestWillBeMoved.getFileName()), REPLACE_EXISTING);

        }

    }
}
