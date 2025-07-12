import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelFormatConverter {
    public static void main(String[] args) {
        String xlsmFilePath = "your_excel_file.xlsm";
        String xlsxFilePath = "converted_file.xlsx";

        try (FileInputStream fis = new FileInputStream(xlsmFilePath);
             Workbook workbook = XSSFWorkbookFactory.create(fis, null);
             FileOutputStream fos = new FileOutputStream(xlsxFilePath)) {

            // 直接写入为XLSX格式
            workbook.write(fos);
            System.out.println("转换完成，文件已保存为: " + xlsxFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}    