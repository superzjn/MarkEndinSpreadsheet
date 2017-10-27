
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.stream.Stream;


public class MarkSpreadsheet {


    public static void main(String[] args) throws IOException {
        // create a new file input stream with the input file specified
        // at the command line
        String filePathofSheet = "/Users/jzhang9/Downloads/Store Inventory for usproductdeal.xlsx";
//        String filePathofSheet = "/Users/jzhang9/Downloads/test.xlsx";
        String filePathforEndedIds = "/Users/jzhang9/Downloads/ids.txt";

        FileInputStream fin = new FileInputStream(filePathofSheet);
        // create a new org.apache.poi.poifs.filesystem.Filesystem

        // For xls file
        //POIFSFileSystem poifs = new POIFSFileSystem(fin);
        //HSSFWorkbook wb = new HSSFWorkbook(poifs);

        // For xlsx file
        XSSFWorkbook wb = new XSSFWorkbook(fin);

        DataFormatter formatter = new DataFormatter();
        Sheet sheet1 = wb.getSheet("Active");
        //Sheet sheet1 = wb.getSheetAt(0);

        String ids = loadTxtFile(filePathforEndedIds);

        for (Row row : sheet1) {
            //   for (Cell cell : row) {
//                CellReference cellRef = new CellReference(row.getRowNum(), 0);
//                System.out.print(cellRef.formatAsString());

            Cell cell = row.getCell(0);

            String ebaylink = formatter.formatCellValue(cell);

//            String ebaylink = row.getCell(0).getRichStringCellValue().getString().trim();
            String itemid = ebaylink.substring(ebaylink.lastIndexOf("/") + 1, ebaylink.length());

            System.out.println(itemid);

            cell = row.getCell(4);

            if ((itemid != "") && (ids.contains(itemid))) {
                if (cell != null) {
                    cell.setCellValue("Ended");
                } else {
                    row.createCell(4).setCellValue("Ended");
                }
            }
        }
        // Write the output to a file
        fin.close();
        FileOutputStream fileOut = new FileOutputStream(filePathofSheet);
        wb.write(fileOut);
        wb.close();
        fileOut.close();


    }


    public static String loadTxtFile(String filePath) {

        String ids = "";

        try {
            Stream<String> lines = Files.lines(Paths.get(filePath));
            StringBuilder data = new StringBuilder();
            lines.forEach(line -> data.append(line).append("\n"));
            ids = data.toString();
            lines.close();

        } catch (Exception e) {
            System.out.println(e.getStackTrace());
        }

        if (ids != "") {
            System.out.println("Txt File Loaded");
        }
        return ids;
    }
}



