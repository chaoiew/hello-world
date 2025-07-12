import java.time.LocalDate;
import java.time.Period;

public class AgeCalculator {
    // 根据最近一次生日计算年龄
    public static int calculateAgeByLastBirthday(LocalDate birthDate, LocalDate currentDate) {
        // 计算当前年龄（周岁）
        int currentAge = Period.between(birthDate, currentDate).getYears();
        
        // 计算今年的生日日期
        LocalDate thisYearBirthday = birthDate.withYear(currentDate.getYear());
        
        // 如果今年的生日还未过，则年龄为周岁减1
        if (thisYearBirthday.isAfter(currentDate)) {
            return currentAge - 1;
        }
        
        // 否则，年龄为周岁
        return currentAge;
    }

    public static void main(String[] args) {
        // 示例：出生日期为2000-01-01，当前日期为2023-06-15
        LocalDate birthDate = LocalDate.of(2000, 1, 1);
        LocalDate currentDate = LocalDate.now();
        
        int age = calculateAgeByLastBirthday(birthDate, currentDate);
        System.out.println("根据最近一次生日计算的年龄：" + age);
    }
}
