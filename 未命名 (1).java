import java.time.LocalDate;
import java.time.Period;
import java.time.format.DateTimeFormatter;

public class AgeCalculator {
    // 根据下次生日计算年龄
    public static int calculateAgeByNextBirthday(LocalDate birthDate, LocalDate currentDate) {
        // 计算当前年龄（周岁）
        int currentAge = Period.between(birthDate, currentDate).getYears();
        
        // 计算下次生日的日期
        LocalDate nextBirthday = birthDate.withYear(currentDate.getYear());
        if (nextBirthday.isBefore(currentDate) || nextBirthday.isEqual(currentDate)) {
            nextBirthday = nextBirthday.plusYears(1); // 如果今年的生日已过或今天是生日，下次生日为明年
        }
        
        // 判断是否已满一岁
        Period period = Period.between(currentDate, nextBirthday);
        boolean hasPassedOneYear = period.getYears() > 0 || period.getMonths() > 0 || period.getDays() > 0;
        
        // 如果下次生日在一年后，年龄+1
        return hasPassedOneYear ? currentAge + 1 : currentAge;
    }

    public static void main(String[] args) {
        // 示例：出生日期为2000-01-01，当前日期为2023-06-15
        LocalDate birthDate = LocalDate.of(2000, 1, 1);
        LocalDate currentDate = LocalDate.now();
        
        int age = calculateAgeByNextBirthday(birthDate, currentDate);
        System.out.println("根据下次生日计算的年龄：" + age);
    }
}
