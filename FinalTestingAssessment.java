import java.awt.List;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Sleeper;

public class FinalTestingAssessment {

	public static void main(String args[]) throws FileNotFoundException, IOException
	{
		System.setProperty("webdriver.chrome.driver","C:\\geckodriver\\chromedriver.exe");
		WebDriver driver =new ChromeDriver();
		driver.get("https://www.amazon.in");
		String[] arr={"red label 1 kg","pressure cooker","mens footwear","redmi note 8"};

		String parentWinHandle = driver.getWindowHandle();
		for(int i=0;i<4;i++)
		{
			driver.switchTo().window(parentWinHandle);
		WebElement search=driver.findElement(By.id("twotabsearchtextbox"));
		search.clear();
		search.sendKeys(arr[i]);
		WebElement searchbtn=driver.findElement(By.xpath("/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[2]/div/input"));
		searchbtn.click();
		//WebElement product1=driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/div[2]/div/span[4]/div[1]/div[1]/div/span/div/div/div/div/div[2]/h2/a/span"));
		WebElement product1=driver.findElement(By.className("s-image"));
		product1.click();
		Set<String> winHandles = driver.getWindowHandles();   
		// Loop through all handles
	        for(String handle: winHandles){
	            if(!handle.equals(parentWinHandle)){
	            driver.switchTo().window(handle);
	            WebElement add2cart=driver.findElement(By.id("add-to-cart-button"));
	    		add2cart.click();
	    		//WebElement check=driver.findElement(By.className("a-button-text"));
	    		driver.close();
	            }
	        }
		}
		driver.switchTo().window(parentWinHandle);
		WebElement cart=driver.findElement(By.id("nav-cart"));
		cart.click();
		 java.util.List<WebElement> allitems = driver.findElements(By.className("sc-product-title"));
         int RowCount = allitems.size();
         java.util.List<WebElement> allprice = driver.findElements(By.className("sc-product-price"));
         ArrayList<String> products = new ArrayList<String>();
         ArrayList<String> price =  new ArrayList<String>();
        for(int i=0;i<RowCount;i++)
        {
        	WebElement w1=allitems.get(i);
        	WebElement w2=allprice.get(i);
        	products.add(w1.getText());
        	price.add(w2.getText());
        	System.out.println(w1.getText());
        	System.out.println(w2.getText());
        }
        WebElement total = driver.findElement(By.className("sc-price"));
        System.out.println(total.getText());
		products.add("total");
		price.add(total.getText());
        
        //**********************************************adding to excel sheet******************
		
		 XSSFWorkbook workbook = new XSSFWorkbook();
	        XSSFSheet sheet = workbook.createSheet("products");
	         for(int i=0;i<RowCount+1;i++)
	         {
	        	 Row row = sheet.createRow(i+1);
	        	 for(int j=1;j<=2;j++)
	        	 {
	        		 Cell cell = row.createCell(j);
	        		 if(j==1)
	        		 cell.setCellValue(products.get(i));
	        		 if(j==2)
	        			 cell.setCellValue(price.get(i));
	        	 }
	         }
	         
	        try (FileOutputStream outputStream = new FileOutputStream("D://products.xlsx")) {
	            workbook.write(outputStream);
	        }
	}
}
