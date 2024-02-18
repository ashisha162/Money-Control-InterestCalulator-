package MoneyControl;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class MoneyDrivenTest
{

    public static <XSSRow> void main(String args[]) throws IOException {
        WebDriver driver=new ChromeDriver();
        driver.get("https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india/fixed-deposit-calculator-SBI-BSB001.html");

        FileInputStream file=new FileInputStream("D://DATADRIVEN//Book2.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(file);
        XSSFSheet sheet=workbook.getSheet("Sheet2");

        //using this we will coutn the number of row in the sheet
        int rowcount=sheet.getLastRowNum();

        for(int i=1; i<=rowcount; i++)
        {
            XSSFRow row=sheet.getRow(i);


            //2ND METHOD

            XSSFCell principle =row.getCell(0); // this will retutn the cell object
            int princ=(int)principle.getNumericCellValue();

            XSSFCell ROI=row.getCell(1);
            int rateofInterest =(int)ROI.getNumericCellValue();

            XSSFCell period =row.getCell(2);
            int per=(int)period.getNumericCellValue();

            XSSFCell Freq=row.getCell(3);
            String freq=Freq.getStringCellValue();

            XSSFCell valueMaturity=row.getCell(4);
            int Exp_Value=(int)valueMaturity.getNumericCellValue();


            driver.findElement(By.xpath("//*[@id=\"principal\"]")).sendKeys(String.valueOf(princ));
            driver.findElement(By.xpath("//*[@id=\"interest\"]")).sendKeys(String.valueOf(rateofInterest));
  driver.findElement(By.id("tenure")).sendKeys(String.valueOf(per));

  Select periodcombo=new Select(driver.findElement(By.id("tenurePeriod")));
  periodcombo.selectByVisibleText("year(s)");


  Select frequency=new Select(driver.findElement(By.id("frequency")));
  frequency.selectByVisibleText(freq);

  driver.findElement(By.xpath("//*[@id=\"fdMatVal\"]/div[2]/a[1]/img")).click();


  String actual_mvalue=driver.findElement(By.xpath("//*[@id=\"resp_matval\"]")).getText();

  if(Double.parseDouble(actual_mvalue)==Exp_Value)
  {
      System.out.println("Test Passed");
  }
  else {
      System.out.println("Test case fail");
  }
driver.findElement(By.xpath("//*[@id=\"fdMatVal\"]/div[2]/a[2]/img")).click();

        }

driver.close();
        driver.quit();

    }
}
