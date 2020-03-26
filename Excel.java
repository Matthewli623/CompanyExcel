import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class Excel {

    private static final String FILE_NAME = "PP 201811Nov r40000-80000.xlsx";

    private static List<String> brandList = new ArrayList<>();

    public static FormulaEvaluator evaluator;

    public static int chartoint(char input) {
        int num = input;
        return num - 65;
    }

    public static String searchBrand(String input) {
        int start = 0;
        int end = 0;
        String res = "";
        for (int i = 0; i < brandList.size(); i++) {
            if (input.contains(brandList.get(i))) {
                return brandList.get(i);
            }
        }
        try {
            for (int i = 0; i < input.length(); i++) {

                int codepoint = input.codePointAt(i);
                if (Character.UnicodeScript.of(codepoint) != Character.UnicodeScript.HAN) {
                    if (Character.isLetter(input.charAt(i))) {
                        int j = i + 1;
                        start = i;
                        int len = 1;
                        while (j < input.length()) {
                            if (Character.isLetter(input.charAt(j)) || Character.isSpaceChar(input.charAt(j))
                                    || input.charAt(j) == '+') {
                                j++;
                                len++;
                            } else {
                                break;
                            }
                        }
                        if (j >= input.length())
                            j--;
                        if (input.charAt(j - 1) == '牌')
                            j--;

                        if (len > 3) {

                            return input.substring(i, j); // j = index+1
                        }
                    }
                }
            }
            return "无品牌";
        } catch (Exception e) {
            // TODO: handle exception
            return "无品牌";
        }

    }

    public static String searchModel(String input) {
        int start = 0;
        int end = 0;
        String res = "";

        for (int i = 0; i < input.length(); i++) {
            int codepoint = input.codePointAt(i);
            if (Character.UnicodeScript.of(codepoint) != Character.UnicodeScript.HAN) {
                if (Character.isLetter(input.charAt(i)) || Character.isDigit(input.charAt(i)))// start with char or
                                                                                              // digit
                {
                    int j = i + 1;
                    start = i;
                    int len = 1;
                    while (j < input.length()) {
                        if ((Character.isLetter(input.charAt(j))) || Character.isDigit(input.charAt(j))
                                || input.charAt(j) == '-')

                        {
                            j++;
                            len++;
                        } else {
                            break;
                        }
                    }
                    // if(j>input.length())j--;
                    // if(input.charAt(j-1)=='牌')j--;
                    int f_charcounter = 0;
                    boolean isSpec = false;
                    for (int k = i; k < j; k++) {
                        codepoint = input.codePointAt(k);
                        if (Character.UnicodeScript.of(codepoint) == Character.UnicodeScript.HAN)// when end with
                                                                                                 // chinese char
                        {
                            j = k;
                            break;
                        }
                        if (input.charAt(k) == 'V' || input.charAt(k) == 'W')
                            f_charcounter++;

                    }
                    if (f_charcounter == 1)
                        isSpec = true;
                    res = input.substring(i, j);
                    if (len >= 4 && j - i > 3 && res.matches(".*\\d.*") && !isSpec)// contain digit
                    {

                        return res; // j = index+1
                    }
                }
            }
        }
        return "无型号";
    }

    public static void main(String[] args) {

        brandList.add("无品牌");
        try {
            System.out.println("Opening excel");
            FileInputStream fis = new FileInputStream(FILE_NAME);
            Workbook wb = new XSSFWorkbook(fis);
            int sheetIndex = 0;
            Sheet sheet = wb.getSheetAt(sheetIndex);

            // 42284
            int startingrow = 3;
            for (int i = startingrow; i < sheet.getPhysicalNumberOfRows() - 1; i++) {
                System.out.println("processing " + i);
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(chartoint('K'));// char of Name
                String brand = "";
                String model = "";
                switch (cell.getCellType()) {
                    case STRING:
                        brand = searchBrand(cell.getStringCellValue());
                        model = searchModel(cell.getStringCellValue());
                        System.out.println(brand);
                        System.out.println(model);
                        break;
                    default:
                        brand = "无品牌";
                }

                // System.out.println(brand);
                // A=0 B = 1 C=2
                try {
                    Cell brandCell = row.getCell(chartoint('M'));
                    if (brandCell == null) {
                        brandCell = sheet.getRow(i).createCell(chartoint('M'));
                    }
                    brandCell.setCellValue(brand);

                    Cell modelCell = row.getCell(chartoint('L'));
                    if (modelCell == null) {
                        modelCell = sheet.getRow(i).createCell(chartoint('L'));
                    }
                    modelCell.setCellValue(model);
                } catch (Exception e) {
                    e.printStackTrace();
                }

            }

            fis.close();
            System.out.println("Writing data");
            FileOutputStream outFile = new FileOutputStream(FILE_NAME);
            wb.write(outFile);
            outFile.close();
            System.out.println("Finished");
        } catch (IOException e) {

        }

    }

}