import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;

import static java.sql.DriverManager.*;

public class ExcelToDB {
    public static void main(String[] args)  {

        //Database connection
        Connection connection = null;
        //URL к базе состоит из протокола:подпротокола://[хоста]:[порта_СУБД]/[БД] и других_сведений
        String url = "jdbc:postgresql://127.0.0.1:5432/northwind";
        //Имя пользователя БД
        String name = "postgres";
        //Пароль
        String password = "1965";
        try {
            //Загружаем драйвер
            Class.forName("org.postgresql.Driver");
            System.out.println("Драйвер подключен");
            //Создаём соединение
            connection = DriverManager.getConnection(url, name, password);
            System.out.println("Соединение установлено");

            //Connection con= getConnection("jdbc:postgresql://127.0.0.1:5432/northwind","postgres","1965");
            Statement stmt=connection.createStatement();

            //create a new table in the database
            String sql ="drop table testtable; create table testtable (test1 varchar(100),test2 varchar(100),test3 varchar(100),test4 varchar(100),test5 varchar(100),test6 varchar(100));";
            stmt.execute(sql);

            //Excel
            FileInputStream fis = new FileInputStream("/home/ivanov/IdeaProjects/html-exporter-master/report.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheet("Sheet0");

            int rows = sheet.getLastRowNum();

            for (int r = 1; r <= rows; r++) {
                XSSFRow row = sheet.getRow(r);
                String test1 = row.getCell(0).getStringCellValue();
                String test2 = row.getCell(1).getStringCellValue();
                String test3 = row.getCell(2).getStringCellValue();
                String test4 = row.getCell(3).getStringCellValue();
                String test5 = row.getCell(4).getStringCellValue();
                String test6 = row.getCell(5).getStringCellValue();

                sql = "insert into testtable values('" + test1 + "', '" + test2 + "', '" + test3 + "', '" + test4 + "', '" + test5 + "', '" + test6 + "')";
                stmt.execute(sql);
                stmt.execute("commit");
            }


            workbook.close();
            fis.close();
            connection.close();

            System.out.println("Done!!");


        } catch (Exception ex) {
            //выводим наиболее значимые сообщения
            Logger.getLogger(ExcelToDB.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            if (connection != null) {
                try {
                    connection.close();
                } catch (SQLException ex) {
                    Logger.getLogger(ExcelToDB.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }

    }
}
