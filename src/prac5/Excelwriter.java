package prac5;

import jxl.*;
import jxl.write.*;
import jxl.write.Number;
import java.io.*;
import java.util.Locale;
public class Excelwriter {
    public static void main(String[] args) throws IOException, WriteException {
        int r = 0, c = 0;
        String header[] = {"Student Name", "Sub1", "Sub2", "Sub3", "Total"};
        String sname[] = {"Sahil", "Anurag", "Siddesh", "Sudhesh", "Krishna", "Piyush", "Rohit", "Vighnesh", "Rohan", "Shyam"};
        int marks[] = {70, 72, 78, 99, 76, 95, 83, 90, 71, 87};
        File file = new File("student.xls");
        WorkbookSettings wbsettings = new WorkbookSettings();
        wbsettings.setLocale(new Locale("en", "EN"));
        WritableWorkbook workbook = Workbook.createWorkbook(file, wbsettings);
        workbook.createSheet("Report", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        for (r = 0; r < 1; r++) {
            for (c = 0; c < header.length; c++) {
                Label l = new Label(c, r, header[c]);
                excelSheet.addCell(l);
            }
        }
        for (r = 1; r <= sname.length; r++) {
            for (c = 0; c < 1; c++) {
                Label l = new Label(c, r, sname[r - 1]);
                excelSheet.addCell(l);
            }
        }
        for (r = 1; r <= marks.length; r++) {
            for (c = 1; c < 4; c++) {
                Number num = new Number(c, r, marks[r - 1]);
                excelSheet.addCell(num);
            }
        }
        for (r = 1; r <= sname.length; r++) {
            for (c = 4; c < 5; c++) {
                int total = marks[r - 1] * 3;  
                Number num = new Number(c, r, total);
                excelSheet.addCell(num);
            }
        }
        workbook.write();
        workbook.close();
        System.out.println("Excel file created successfully!");
    }
}
