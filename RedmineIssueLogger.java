/*
REDMINE ISSUE LOGGER  (22-03-2018) Version 1.9 [Updated on 28th November 2018]
--> Developed by B B Susheel Kumar (susheelkumar@gmobis.com)
This script can be used to change the status of the issue reported in Redmine by giving 
the Excel Sheet in the format specified as input to the script. Refer: IssuesFormat.xlsx
Developed using the Java Libraries of Apache POI and Selenium using Chrome Driver. 
Works for Chrome and can be extended to other browsers using respective selenium browser drivers.
*/
 
import java.awt.Container;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.PrintStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.ClickAction;
 
 
@SuppressWarnings({ "unused", "deprecation" })
public class RedmineIssueLogger {
 
	public static String sUserName = "rama.dabbi@gmobis.com";
	public static String sPassWord = "rama.dabbi";
	public static String fileAddr = "U:/RedmineCloser/RedmineLoggingSheet_iot.xlsx";
	public static XSSFRow row;
	public static Cell currentCell,previousCell;
	static DataFormatter dataFormatter = new DataFormatter();
	static boolean Cflag = true;
	static int count = 0;
	static WebDriver driver = null;
	static String STMS_ID,TASK_TITLE,DESCRIPTION,PRIORITY,REQUEST_TYPE,CATEGORY,SOURCE_REGION,CAR_MODEL,BUILD_VERSION,PROBABILITY,PLATFORM,MODULE,ISSUE_CLASSIFICATION;
 
	public static void main(String args[])
	{
		RedmineIssueLogger issueLogger = new RedmineIssueLogger();	
		try{
 
		Date date = new Date() ;
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss") ;
        PrintStream o = new PrintStream(new File("U:/RedmineCloser/"+dateFormat.format(date) + " RedmineLogger.txt")); 
        System.setOut(o); //ENABLE FOR LOG IN TXT FILE
 
		System.setProperty("webdriver.chrome.driver","U:/RedmineCloser/chromedriver.exe");
		driver = new ChromeDriver();
		driver.navigate().to("http://10.126.161.136:3001/");
		Cflag = false;
		driver.manage().window().maximize();
		driver.findElement(By.xpath("//*[@id=\"account\"]/ul/li[1]/a")).click();
		//BEGIN LOGIN
		WebElement userId = driver.findElement(By.xpath("//*[@id=\"username\"]"));
		userId.click();
		userId.clear();
		userId.sendKeys(new CharSequence[]{sUserName});
		WebElement passWord = driver.findElement(By.xpath("//*[@id=\"password\"]"));
		passWord.click();
		passWord.clear();
		passWord.sendKeys(new CharSequence[]{sPassWord});
		driver.findElement(By.xpath("//*[@id=\"login-submit\"]")).click();
		try{
 
			if(driver.findElement(By.xpath("//*[@id=\"flash_error\"]")) != null){
			throw new Exception("Wrong Password");
			}
		}
		catch(Exception e){
			if(e.getMessage().equalsIgnoreCase("Wrong Password\n")){
				System.out.println("Exit because of wrong password....\n");
				driver.quit();
				return;	
			}
		}
		//END LOGIN
		//BEGIN NAVIGATION TO NEW ISSUE
		driver.findElement(By.xpath("//*[@id=\"top-menu\"]/ul/li[3]/a")).click();
		driver.findElement(By.xpath("//*[@id=\"projects-index\"]/ul/li[1]/ul/li/ul/li[1]/ul/li[2]/ul/li[2]/div/a")).click();
		driver.findElement(By.xpath("//*[@id=\"main-menu\"]/ul/li[4]/a")).click();
		//END NAVIGATION TO NEW ISSUE
		//BEGIN OPENING AN EXCEL FILE
		File file = new File(fileAddr);
		FileInputStream excelFile = new FileInputStream(file);
	    if(file.isFile() && file.exists()) {
	         System.out.println("File open successful.\n");
	      } else {
	         System.out.println("Error to open file.\n");
	         driver.quit();
	         return;
	      }
	    XSSFWorkbook workbook = new XSSFWorkbook(excelFile); 
	    XSSFSheet spreadsheet = workbook.getSheetAt(0);
	    Iterator < Row >  rowIterator = spreadsheet.iterator();
	    rowIterator.next();//first row contains the data about the columns
	    rowIterator.next();
 
	    while(rowIterator.hasNext()){
	    	driver.findElement(By.xpath("//*[@id=\"issue_tracker_id\"]/option[1]")).click();//Tracker:MM-01 ISSUES & FEATURE REQ
	    	row = (XSSFRow)rowIterator.next();
	    	Iterator < Cell >  cellIterator = row.cellIterator();
	    	currentCell = cellIterator.next();
	    	STMS_ID = currentCell.getStringCellValue();
	    	currentCell = cellIterator.next();
	    	TASK_TITLE = currentCell.getStringCellValue();
	    	WebElement subject = driver.findElement(By.xpath("//*[@id=\"issue_subject\"]"));
	    	subject.click();
	    	subject.clear();
	    	subject.sendKeys(STMS_ID+" "+TASK_TITLE);
	    	currentCell = cellIterator.next();
	    	DESCRIPTION = currentCell.getStringCellValue();
	    	WebElement Desc = driver.findElement(By.xpath("//*[@id=\"issue_description\"]"));
	    	Desc.click();
	    	Desc.clear();
	    	Desc.sendKeys(DESCRIPTION);
	    	currentCell = cellIterator.next();
	    	PRIORITY = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_priority_id\"]/option["+issueLogger.getPriorityIndex(PRIORITY)+"]")).click();
	    	//driver.findElement(By.xpath("//*[@id=\"issue_assigned_to_id\"]/option[161]")).click();
	    	currentCell = cellIterator.next();
	    	REQUEST_TYPE = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_123\"]/option["+issueLogger.getRequestIndex(REQUEST_TYPE)+"]")).click();
	    	currentCell = cellIterator.next();
	    	CATEGORY = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_112\"]/option["+issueLogger.getCategory(CATEGORY)+"]")).click();//CATEGORY
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_54\"]/option[3]")).click();//ORIGIN OF REQUEST
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_113\"]/option[2]")).click();//SPECIFY THE SOURCE FROM WHICH ISSUE ID IS COPIED
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_26\"]")).sendKeys(STMS_ID.replaceAll("[^0-9]", ""));
	    	currentCell = cellIterator.next();
	    	SOURCE_REGION = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_50\"]/option["+issueLogger.getSourceRegion(SOURCE_REGION)+"]")).click();//CATEGORY
	    	currentCell = cellIterator.next();
	    	CAR_MODEL = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_33\"]/option["+issueLogger.getCarModel(CAR_MODEL)+"]")).click();//CARMODEL
	    	currentCell = cellIterator.next();
	    	BUILD_VERSION = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_74\"]")).sendKeys(BUILD_VERSION);
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_39\"]")).sendKeys("NA");//MICOM VERSION
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_37\"]")).sendKeys("NA");//BT VERSION
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_56\"]")).sendKeys("NA");//VG VERSION
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_58\"]")).sendKeys("NA");//VR VERSION
	    	currentCell = cellIterator.next();
	    	PROBABILITY = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_53\"]/option["+issueLogger.getProbability(PROBABILITY)+"]")).click();//FREQUENCY
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_62\"]/option[7]")).click();//PLATFORM VARIANTS
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_65\"]/option[9]")).click();//PRELIMINARY CASUAL ANAYLYSIS
	    	currentCell = cellIterator.next();
	    	PLATFORM = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_70\"]/option["+issueLogger.getPlatForm(PLATFORM)+"]")).click();//PLATFORM
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_71\"]/option[12]")).click();//ISSUED MILESTONE
	    	currentCell = cellIterator.next();
	    	MODULE = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_108\"]/option["+issueLogger.getModule(MODULE)+"]")).click();//MODULE
	    	currentCell = cellIterator.next();
	    	ISSUE_CLASSIFICATION = currentCell.getStringCellValue();
	    	driver.findElement(By.xpath("//*[@id=\"issue_custom_field_values_109\"]/option["+issueLogger.getIssueClassification(ISSUE_CLASSIFICATION)+"]")).click();//ISSUE_CLASSIFICATION
	    	driver.findElement(By.xpath("//*[@id=\"issue-form\"]/input[3]")).click();
	    	WebElement feedback1 = driver.findElement(By.xpath("//*[@id=\"content\"]/h2"));
	    	WebElement feedback2 = driver.findElement(By.xpath("//*[@id=\"flash_notice\"]"));
	    	String feedbackData1 = feedback1.getText();
	    	String feedbackData2 = feedback2.getText();
	    	if(feedbackData1.contains("ISSUES & FEATURE REQ") && feedbackData2.contains("created"))
	    	{
	    		System.out.println("Logged the issue with STMS ID:"+STMS_ID+" Successfully...\n");
	    		count++;
	    		driver.findElement(By.xpath("//*[@id=\"main-menu\"]/ul/li[4]/a")).click();
	    	}
	    	else{
	    		driver.quit();
	    		System.out.println("Failed to log "+STMS_ID+" :(\n");
	    	}
 
 
	    }
 
		driver.quit();
		}
		catch(Exception e){
			System.out.println(e);
		}
		finally{
			if(Cflag)
			{
				System.out.println("==========================================\nProcess failed to start...\n Kill ChromeDriver and start again....\n");
			}
			System.out.println("Total No.Of Issues Logged:"+count);
			driver.quit();
		}
 
	}
	public int getIssueClassification(String string){
		if(string.equals("Gui"))
			return 2;
		else return 3;
	}
 
	public int getModule(String string){
		if(string.equals("IOT"))
				return 81;
		else if(string.equals("CAN"))
			return 76;
		else return 95;
 
 
	}
	public int getPriorityIndex(String string){
		if(string.equals("A"))
			return 3;
		else if(string.equals("B"))
			return 2;
		else if(string.equals("C"))
			return 1;
		else if(string.equals("TOP"))
			return 4;
		else return 2;
 
	}
 
	public int getRequestIndex(String string)
	{
		if(string.equals("SOFTWARE ISSUE"))
			return 2;
		else if(string.equals("NEW SPECIFICATION"))
			return 3;
		else if(string.equals("CHANGE IN SPECIFICATION"))
			return 4;
		else if(string.equals("CAUSED BY - FRAMEWORK"))
			return 5;
		else if(string.equals("CAUSED BY - OTHERS"))
			return 9;
		else return 2;	
	}
 
	public int getCategory(String string){
		if(string.equalsIgnoreCase("BLUETOOTH"))
			return 2;
		else if(string.equalsIgnoreCase("BROADCAST"))
			return 3;
		else if(string.equalsIgnoreCase("CAN"))
			return 4;
		else if(string.equalsIgnoreCase("CLUSTER"))
			return 5;
		else if(string.equalsIgnoreCase("CONNECTIVITY"))
			return 6;
		else if(string.equalsIgnoreCase("DSP"))
			return 7;
		else if(string.equalsIgnoreCase("ECALL"))
			return 8;
		else if(string.equalsIgnoreCase("MEDIA"))
			return 9;
		else if(string.equalsIgnoreCase("NAVIGATION"))
			return 10;
		else if(string.equalsIgnoreCase("SETTINGS"))
			return 11;
		else if(string.equalsIgnoreCase("SYSTEM"))
			return 12;
		else if(string.equalsIgnoreCase("TELEMATICS"))
			return 13;
		else return 14;			
 
	}
 
	public int getSourceRegion(String string){
		if(string.equalsIgnoreCase("ALL"))
			return 1;
		else if(string.equalsIgnoreCase("AUSTRALIA"))
			return 2;
		else if(string.equalsIgnoreCase("CANADA"))
			return 3;
		else if(string.equalsIgnoreCase("CHINA"))
			return 4;
		else if(string.equalsIgnoreCase("EUROPE"))
			return 5;
		else if(string.equalsIgnoreCase("GERMANY"))
			return 6;
		else if(string.equalsIgnoreCase("INDIA"))
			return 7;
		else if(string.equalsIgnoreCase("KOREA"))
			return 8;
		else if(string.equalsIgnoreCase("MIDDLE EAST"))
			return 9;
		else if(string.equalsIgnoreCase("USA"))
			return 10;
		else return 11;			
	}
 
	public int getCarModel(String string){
		if(string.toLowerCase().contains("ab"))
			return 1;
		else if(string.toLowerCase().contains("ad"))
			return 2;
		else if(string.toLowerCase().contains("ad pe"))
			return 3;
		else if(string.toLowerCase().contains("adi"))
			return 4;
		else if(string.toLowerCase().contains("adpe"))
			return 5;
		else if(string.toLowerCase().contains("ae"))
			return 6;
		else if(string.toLowerCase().contains("ah"))
			return 7;
		else if(string.toLowerCase().contains("ah2"))
			return 8;
		else if(string.toLowerCase().contains("ai3"))
			return 9;
		else if(string.toLowerCase().contains("bape"))
			return 10;
		else if(string.toLowerCase().contains("bd"))
			return 11;
		else if(string.toLowerCase().contains("bdm"))
			return 12;
		else if(string.toLowerCase().contains("br2"))
			return 13;
		else if(string.toLowerCase().contains("dn8"))
			return 21;
		else if(string.toLowerCase().contains("lx2"))
			return 43;
		else if(string.toLowerCase().contains("qlc"))
			return 59;
		else if(string.toLowerCase().contains("qxi"))
			return 62;
		else if(string.toLowerCase().contains("sk"))
			return 64;
		else if(string.toLowerCase().contains("sk3"))
			return 65;
		else if(string.toLowerCase().contains("tlc"))
			return 70;
		else return 82;
 
	}
 
	public int getProbability(String string){
		if(string.toLowerCase().contains("always"))
			return 3;
		else if(string.toLowerCase().contains("sometimes"))
			return 4;
		else if(string.toLowerCase().contains("rarely"))
			return 5;
		else if(string.toLowerCase().contains("not-reproducible"))
			return 6;
		else return 2;
	}
 
	public int getPlatForm(String string){
		if(string.toLowerCase().contains("gen5"))
			return 8;
		else if(string.toLowerCase().contains("d-audio"))
			return 4;
		else if(string.toLowerCase().contains("premium"))
			return 11;
		else return 15;
 
	}
 
 
}
 
 
