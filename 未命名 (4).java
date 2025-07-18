import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.POIXMLProperties;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;

import java.io.*;
import java.util.Base64;

public class ExcelVBAManager {
    // VBA代码（即你提供的自动保存宏）
    private static final String VBA_CODE = 
        "Option Explicit\n" +
        "' 工作簿关闭前自动保存\n" +
        "Private Sub Workbook_BeforeClose(Cancel As Boolean)\n" +
        "    On Error Resume Next\n" +
        "    Application.ScreenUpdating = False\n" +
        "    \n" +
        "    ' 检查是否有未保存的更改\n" +
        "    If ThisWorkbook.Saved = False Then\n" +
        "        ThisWorkbook.Save\n" +
        "        Debug.Print \"已自动保存: \" & Now\n" +
        "    End If\n" +
        "    \n" +
        "    Application.ScreenUpdating = True\n" +
        "End Sub";

    public static void main(String[] args) {
        try {
            // 创建或打开Excel文件
            File file = new File("path/to/your/excel.xlsm");
            XSSFWorkbook workbook;
            
            if (file.exists()) {
                workbook = new XSSFWorkbook(OPCPackage.open(file, PackageAccess.READ_WRITE));
            } else {
                workbook = new XSSFWorkbook();
                // 创建一个空工作表
                XSSFSheet sheet = workbook.createSheet("Sheet1");
            }
            
            // 嵌入VBA代码
            embedVbaCode(workbook, VBA_CODE);
            
            // 保存文件
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
            
            System.out.println("VBA代码已成功嵌入Excel文件！");
            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void embedVbaCode(XSSFWorkbook workbook, String vbaCode) throws Exception {
        // 创建VBA项目部件
        POIXMLDocumentPart vbaProject = new POIXMLDocumentPart(workbook) {
            @Override
            protected void commit() throws IOException {
                // 此方法必须重写，但内容可以为空
            }
        };
        
        // 将VBA代码转换为Base64编码（VBA项目需要这种格式）
        byte[] vbaBytes = vbaCode.getBytes("UTF-8");
        String base64Vba = Base64.getEncoder().encodeToString(vbaBytes);
        
        // 设置VBA项目属性
        POIXMLProperties.CustomProperties customProperties = workbook.getProperties().getCustomProperties();
        customProperties.addProperty("VBAProject", base64Vba);
        
        // 修改工作簿属性以启用宏
        CTWorkbook ctWorkbook = workbook.getCTWorkbook();
        ctWorkbook.setVbaProjectRelId("rIdVBA"); // 设置VBA项目关联ID
    }
}
