package finalexam.exam;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

import java.io.File;


import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;



import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import org.openqa.selenium.By;

public class NewTest {
	public static ExtentSparkReporter sparkReporter;
	public static ExtentReports extent;
	public static ExtentTest test;
	static Logger logger = Logger.getLogger(NewTest.class);
	WebDriver driver;
	
	
	public void initializer() {
		sparkReporter =  new ExtentSparkReporter(System.getProperty("user.dir")+"/Reports/extentSparkReport.html");
		sparkReporter.config().setDocumentTitle("Automation Report");
		sparkReporter.config().setReportName("Test Execution Report");
		sparkReporter.config().setTheme(Theme.STANDARD);
		sparkReporter.config().setTimeStampFormat("yyyy-MM-dd HH:mm:ss");
		extent = new ExtentReports();
		extent.attachReporter(sparkReporter);		
	}
	  
	public static String captureScreenshot(WebDriver driver) throws IOException {
		String FileSeparator = System.getProperty("file.separator"); // "/" or "\"
		String Extent_report_path = "."+FileSeparator+"Reports"; // . means parent directory
		File Src = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		String Screenshotname = "screenshot"+Math.random()+".png";
		File Dst = new File(Extent_report_path+FileSeparator+"Screenshots"+FileSeparator+Screenshotname);
		FileUtils.copyFile(Src, Dst);
		String absPath = Dst.getAbsolutePath();
		//System.out.println("Absolute path is:"+absPath);
		return absPath;
	}
	
  @Test(priority = 1)
  public void openLink() throws InterruptedException {
	
	  logger.info("Program started");
	  Thread.sleep(3000);
	  driver.findElement(By.xpath("//*[@id='OBLIGATIONS_heading']"));
  }
  
  @Test(priority = 2)
  public void identifyTables() throws InterruptedException, IOException {
	    Thread.sleep(3000);
	    

	    String outputExcelFilePath = "C:\\Users\\Owner\\exam\\src\\test\\Excel\\DriverFileOP.xlsx";
	    XSSFWorkbook workbook = new XSSFWorkbook();
	    FileOutputStream outputStream = new FileOutputStream(outputExcelFilePath);

	    String methodName = new Exception().getStackTrace()[0].getMethodName();
		String className = new Exception().getStackTrace()[0].getClassName();
		
	    for (int tableIndex = 2; tableIndex <= 3; tableIndex++) { // Loop through tables 2 to 6
	        List<WebElement> links = driver.findElements(By.xpath(".//table[contains(@summary, 'Obligation')]/tbody[" + tableIndex + "]/tr/td[starts-with(@headers, 'NOM_OFFCL')]/a"));
	        System.out.println(links.size());
	        if (links.size() == 0) {
	            System.out.println("No links found for tableIndex " + tableIndex);
	            continue; // Skip this index if no links found
	        }
	        // Create a sheet for each link
	        Thread.sleep(3000);
	        for (WebElement link : links) {
	        	  Thread.sleep(3000);
	            String sheetName = link.getText();
	            Thread.sleep(3000);
	            System.out.print(sheetName);
	            Thread.sleep(3000);
	           XSSFSheet sheet = workbook.createSheet(sheetName);
	           Thread.sleep(3000);
	            driver.findElement(By.xpath(".//table[contains(@summary, 'Obligation')]/tbody[" + tableIndex + "]/tr/td[starts-with(@headers, 'NOM_OFFCL')]/a")).click();
	            test = extent.createTest(methodName,"Final Exam");
	    		test.log(Status.PASS, "Clicked on the link");
	    		test.assignCategory("Regression Testing");
	    		test.addScreenCaptureFromPath(captureScreenshot(driver));
	    		logger.info("First Screenshot captured");
	    		Thread.sleep(3000); 
	            driver.switchTo().frame(driver.findElement(By.xpath("//*[@id='apex_dialog_1']/iframe")));
	            test.log(Status.PASS, "Web Table");
	    
	    		test.addScreenCaptureFromPath(captureScreenshot(driver));
	    		logger.info("Second Screenshot captured");
	    		Thread.sleep(3000); 
	            WebElement tableEle = driver.findElement(By.xpath("//*[@id='R1740668184739222315']/div[2]/div[2]/table[2]"));

	            //Thread.sleep(3000); 
	            List<WebElement> rows = tableEle.findElements(By.tagName("tr"));

	            //Thread.sleep(3000); 
	            for (int i = 1; i < rows.size(); i++) {
		            WebElement row = rows.get(i);
		            // Find all columns in the row
		            List<WebElement> columns = row.findElements(By.tagName("td"));
		            System.out.print("No. of columns are: "+columns.size() + "\t");
		            // Iterate through columns and print data
		            XSSFRow excelRow = sheet.createRow(i);
                    for (int j = 0; j < columns.size(); j++) {
                        WebElement column = columns.get(j);

        	            Thread.sleep(3000); 
                        String cellData = column.getText();
        
                        // Write cell data to the Excel cell
                        excelRow.createCell(j).setCellValue(cellData);
                    }
                    //driver.switchTo().defaultContent();
                   
                }
	            Thread.sleep(1000);
	            logger.info("Wrote data into excel");
	      	 
	            Thread.sleep(3000); 
	        	}
	         
	        driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
	        Thread.sleep(3000);
            links = driver.findElements(By.xpath(".//table[contains(@summary, 'Obligation')]/tbody[" + tableIndex + "]/tr/td[starts-with(@headers, 'NOM_OFFCL')]/a")); // Refresh links
	        }
	    

	    
	    workbook.write(outputStream);

        Thread.sleep(3000); 
	   
	    outputStream.close();
	    workbook.close();
	}
  
  
  @AfterTest
	public void closeMethod() {
		extent.flush();
		driver.quit();
	}
	
	@BeforeTest
	public void driverSetup() {
		PropertyConfigurator.configure("src\\test\\Excel\\log4j.properties");
		initializer();
		driver = new ChromeDriver();
		driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
	}
	
} 
			
