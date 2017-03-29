import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.write.Label;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Random;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello World!");
        try {
            String fileName = "math_boy" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xls";
            FileOutputStream fos = new FileOutputStream("d:\\" + fileName);
            //创建工作薄
            WritableWorkbook workbook = Workbook.createWorkbook(fos);
            //创建新的一页
            for (int j = 0; j < 3; j++) {
                WritableSheet sheet = workbook.createSheet("1-20加减法"+j,j);
                jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.TIMES, 28, WritableFont.BOLD, true);
                jxl.write.WritableCellFormat wcfF = new jxl.write.WritableCellFormat(wf);
                for (int i = 0; i < 14; i++) {
                    Label label1 = new Label(0, i, getMathProblem() ,wcfF);
                    Label label2 = new Label(1, i, getMathProblem() ,wcfF);
                    Label label3 = new Label(2, i, getMathProblem() ,wcfF);

                    sheet.addCell(label1);
                    sheet.addCell(label2);
                    sheet.addCell(label3);
                    sheet.setColumnView(i,28);
                    sheet.setRowView(i,1000);
                }
            }
            workbook.write();
            workbook.close();
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String getMathProblem() {
        int number1 = new Random().nextInt(15);
        int number2 = new Random().nextInt(10);
        int type = new Random().nextInt(2);
        String content = "";
        switch (type) {
            case 0:
                content = number1 + " + " + number2 + " =   ";
                break;
            case 1:
                content = number1 + " - " + number2 + " =   ";
                break;
            case 2:
                content = number1 + " * " + number2 + " =   ";
                break;
            case 3:
                content = number1 + " ÷ " + number2 + " =   ";
                break;
        }
        return content;

    }
}
