import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class ExcelDesktopOpener {
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\YourName\\Desktop\\example.xlsx"; // 替换为实际Excel文件路径
        openAndCloseExcel(excelFilePath, 10); // 打开Excel，10秒后关闭
    }

    public static void openAndCloseExcel(String filePath, int waitSeconds) {
        try {
            // 打开Excel文件
            Process excelProcess = openExcel(filePath);
            System.out.println("Excel已打开，文件路径: " + filePath);
            
            // 等待指定时间
            System.out.println("等待 " + waitSeconds + " 秒后关闭Excel...");
            TimeUnit.SECONDS.sleep(waitSeconds);
            
            // 关闭Excel进程
            closeExcel(excelProcess);
            System.out.println("Excel已关闭");
            
        } catch (IOException | InterruptedException e) {
            System.err.println("操作Excel时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static Process openExcel(String filePath) throws IOException {
        String os = System.getProperty("os.name").toLowerCase();
        Process process;
        
        if (os.contains("win")) {
            // Windows系统：使用start命令
            process = Runtime.getRuntime().exec("cmd /c start excel.exe " + filePath);
        } else if (os.contains("mac")) {
            // macOS系统：使用open命令
            process = Runtime.getRuntime().exec("open -a \"Microsoft Excel\" " + filePath);
        } else if (os.contains("nix") || os.contains("nux")) {
            // Linux系统：使用libreoffice命令
            process = Runtime.getRuntime().exec("libreoffice --calc " + filePath);
        } else {
            throw new UnsupportedOperationException("不支持的操作系统: " + os);
        }
        
        return process;
    }

    private static void closeExcel(Process excelProcess) throws IOException {
        String os = System.getProperty("os.name").toLowerCase();
        
        if (os.contains("win")) {
            // Windows系统：使用taskkill命令结束excel进程
            Runtime.getRuntime().exec("taskkill /f /im excel.exe");
        } else if (os.contains("mac") || os.contains("nix") || os.contains("nux")) {
            // macOS/Linux系统：使用pkill命令结束excel进程
            Runtime.getRuntime().exec("pkill -f excel");
        } else {
            // 尝试通过Java进程对象销毁（可能无效）
            excelProcess.destroy();
        }
    }
}
