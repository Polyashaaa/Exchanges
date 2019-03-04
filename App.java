import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Locale;

public class App {

    public static void main(String[] args) throws IOException {

        FileInputStream fis = new FileInputStream("C:/needed/TRD.xls");
        Workbook wb = new HSSFWorkbook(fis);
        int max = 0;
        int i = 0;
        int numberOfStart = 0;
        int numberOfEnd = 0;
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("HH:mm:ss.SSS").withLocale(Locale.US);

        while (wb.getSheetAt(0).getRow(i)!=null){
            LocalTime time = LocalTime.parse(wb.getSheetAt(0).getRow(i).getCell(0).getStringCellValue(), formatter);
            LocalTime endingTime = time.plusNanos(999999999);
            int count = 0;
            int j = i;
            while (time.isBefore(endingTime) & wb.getSheetAt(0).getRow(j+1)!=null) {
                time = LocalTime.parse(wb.getSheetAt(0).getRow(j+1).getCell(0).getStringCellValue(), formatter);
                count++;
                j++;
            }

            if (count > max) {
                max = count;
                numberOfStart = i;
                numberOfEnd = j-1;
            }

            i++;
        }

        LocalTime a = LocalTime.parse(wb.getSheetAt(0).getRow(numberOfStart).getCell(0).getStringCellValue(), formatter);
        LocalTime b = LocalTime.parse(wb.getSheetAt(0).getRow(numberOfEnd).getCell(0).getStringCellValue(), formatter);

        System.out.println("Максимальное количество сделок в течение одной секунды было между " + a + " и " + b + ". В этот промежуток произошло " + max + " сделок.");

        ArrayList<String> exchanges = new ArrayList<String>();

        i = 0;

        while (wb.getSheetAt(0).getRow(i)!=null) {
            if (exchanges.contains(wb.getSheetAt(0).getRow(i).getCell(3).getStringCellValue())){
            } else {
                exchanges.add(wb.getSheetAt(0).getRow(i).getCell(3).getStringCellValue());
            }
            i++;
        }

        for (int n = 0; n < exchanges.size(); n++)  {
            max = 0;
            i = 0;
            numberOfStart = 0;
            numberOfEnd = 0;
            while (wb.getSheetAt(0).getRow(i)!=null){
                LocalTime time = LocalTime.parse(wb.getSheetAt(0).getRow(i).getCell(0).getStringCellValue(), formatter);
                LocalTime endingTime = time.plusNanos(999999999);
                int count = 0;
                int j = i;
                while (time.isBefore(endingTime) & wb.getSheetAt(0).getRow(j+1)!=null) {
                    time = LocalTime.parse(wb.getSheetAt(0).getRow(j+1).getCell(0).getStringCellValue(), formatter);
                    if(wb.getSheetAt(0).getRow(i).getCell(3).getStringCellValue().equals(exchanges.get(n))){
                        count++;
                    }
                    j++;
                }

                if (count > max) {
                    max = count;
                    numberOfStart = i;
                    numberOfEnd = j-1;
                }

                i++;
            }

            a = LocalTime.parse(wb.getSheetAt(0).getRow(numberOfStart).getCell(0).getStringCellValue(), formatter);
            b = LocalTime.parse(wb.getSheetAt(0).getRow(numberOfEnd).getCell(0).getStringCellValue(), formatter);

            System.out.println("Максимальное количество сделок в течение одной секунды на бирже " + exchanges.get(n)+ " было между " + a + " и " + b + ". В этот промежуток произошло " + max + " сделок.");
        }
        fis.close();
    }
}
