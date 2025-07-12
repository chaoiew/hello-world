import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellReference;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class ExcelFormulaProcessor {
    public static void main(String[] args) {
        String filePath = "your_excel_file.xlsx";
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 给指定sheet的行列赋值
            Sheet dataSheet = workbook.getSheet("DataSheet");
            Row row = dataSheet.getRow(0);
            if (row == null) {
                row = dataSheet.createRow(0);
            }
            Cell cell = row.getCell(0);
            if (cell == null) {
                cell = row.createCell(0);
            }
            cell.setCellValue(100);

            // 设置另一个单元格的值
            cell = row.getCell(1);
            if (cell == null) {
                cell = row.createCell(1);
            }
            cell.setCellValue(200);

            // 强制计算公式
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.setIgnoreMissingWorkbooks(true);
            evaluator.evaluateAll();

            // 获取用户输入的单元格引用
            Scanner scanner = new Scanner(System.in);
            System.out.print("请输入要读取的单元格坐标（例如E3）：");
            String cellReferenceStr = scanner.nextLine();
            scanner.close();

            // 解析单元格引用
            CellReference cellReference = new CellReference(cellReferenceStr);
            int rowIndex = cellReference.getRow();
            int columnIndex = cellReference.getCol();

            // 读取结果sheet的值
            Sheet resultSheet = workbook.getSheet("ResultSheet");
            Row resultRow = resultSheet.getRow(rowIndex);
            if (resultRow == null) {
                System.out.println("错误：行 " + (rowIndex + 1) + " 不存在");
                return;
            }
            Cell resultCell = resultRow.getCell(columnIndex);
            if (resultCell == null) {
                System.out.println("错误：单元格 " + cellReferenceStr + " 不存在");
                return;
            }

            // 根据单元格类型获取值
            switch (resultCell.getCellType()) {
                case STRING:
                    System.out.println(cellReferenceStr + " 的值: " + resultCell.getStringCellValue());
                    break;
                case NUMERIC:
                    System.out.println(cellReferenceStr + " 的值: " + resultCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    System.out.println(cellReferenceStr + " 的值: " + resultCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    // 已经计算过，不应该进入此分支
                    System.out.println(cellReferenceStr + " 的值: " + evaluator.evaluate(resultCell).formatAsString());
                    break;
                default:
                    System.out.println(cellReferenceStr + " 的值: 未知类型");
            }

            // 保存修改
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}    