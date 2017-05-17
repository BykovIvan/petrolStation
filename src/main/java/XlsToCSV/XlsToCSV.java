package XlsToCSV;

import au.com.bytecode.opencsv.CSVWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;


//import java.io.File;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * Created by ahdpe on 23.04.2017.
 */
public class XlsToCSV {

    public static void main(String[] args) {
        ArrayList<String> list = XlsToCSV.readXls("Itmo.xls");
        try {
            XlsToCSV.WriteToCSV(list);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static ArrayList<String> readXls(String fileName) {
        ArrayList<String> result = new ArrayList<String>();
        HSSFWorkbook ExcelWorkBook = null;
        try {
            InputStream inputStream = new FileInputStream(fileName);
            ExcelWorkBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheet = ExcelWorkBook.getSheetAt(0);
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row row = iterator.next();
            Iterator<Cell> iteratorCells = row.iterator();
            while (iteratorCells.hasNext()) {
                Cell cell = iteratorCells.next();
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        result.add(cell.getStringCellValue() + "");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        String strNumeric =  cell.getNumericCellValue() + "";
                        result.add(strNumeric);
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        String strFormula = cell.getNumericCellValue() + "";
                        result.add(strFormula);
                        break;
                    default:
                        break;
                }
            }
            result.add("\n");
        }
    return result;
    }
    public static void WriteToCSV (ArrayList<String> list) throws IOException {
        CSVWriter writer = new CSVWriter(new FileWriter("Result.csv"));
        String[] resultList = list.toArray(new String[list.size()]);
        writer.writeNext(resultList);
        writer.close();
    }
}
