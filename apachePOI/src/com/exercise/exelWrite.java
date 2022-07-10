package com.exercise;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class exelWrite {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("EmployeeData");
        Map<String,Object[]> data = new TreeMap<String,Object[]>();
        data.put("1",new Object[]{"NAME","LAST NAME","EMAIL","PASSWORD","COMPANY","ADRESS","CUTY","ZIP CODE","MOBILE PHONE"});
        data.put("2",new Object[]{"MORONI","VILLALOBOS","MORONIV@HEXAWARE.COM","123","HEXAWARE","AGUASCALIENTES","AGUASCALIENTES","20280","333123123123"});

        Set<String> keyset=data.keySet();
        int rownum=0;
        for (String key: keyset) {
            Row row=sheet.createRow(rownum++);
            Object[] objarr=data.get(key);
            int cellnum=0;
            for (Object obj:objarr) {
                Cell cell= row.createCell(cellnum++);
                if(obj instanceof  String){
                    cell.setCellValue((String)obj);
                }
                else{
                    cell.setCellValue((Integer)obj);
                }
            }
        }
        try {
            FileOutputStream out = new FileOutputStream(new File("apachePoidemo.xlsx"));
            workbook.write(out);
            out.close();
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }
}
