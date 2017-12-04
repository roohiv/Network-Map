package Scanr;
import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import bsh.ParseException;

public class Scanr_NM {

	public static WebDriver driver;
	public static WebDriverWait wait;
	HashMap<String,String> moduleStatusMap;
	HashMap<String,String> moduleIdMap;
	static public String outReportingPath;
	static public String outputPath; 


	int Result_Index,Actual_Table_Data_index,Actual_Table_Data_Val_index;
	static int duplicate = 1;
	String image_Name;
	String exceptionInMethod;

	public void Read(String ExcelPath) throws IOException,InterruptedException {


		File excel = new File(ExcelPath);
		outReportingPath  = ExcelPath.substring(0,ExcelPath.lastIndexOf('\\')+1);
		outputPath=outReportingPath+ExcelPath.substring(ExcelPath.lastIndexOf('\\')+1,ExcelPath.lastIndexOf('.'));
		// Creating instance of Excel Sheet and WorkBook
		FileInputStream fin = new FileInputStream(excel);
		XSSFWorkbook wb = new XSSFWorkbook(fin);
		
	
		
		
		//Instance of DataSheet
		XSSFSheet ws = wb.getSheetAt(0);

		XSSFRow row=ws.getRow(0);
		Iterator<Cell> iterator=row.cellIterator();

		int i=0,Url_Index=0,TestCase_Name_Index=0,User_ID_Index=0,Password_Index=0,BAN_Index=0,Table_Data_Index=0,Device_Index=0,Table_Data_Val_Index=0,Execute_Index=0;

		while(iterator.hasNext())
		{ Cell cell=iterator.next();
		if(cell.getStringCellValue().equalsIgnoreCase("Url")){
			Url_Index=i;
		}else if(cell.getStringCellValue().contains("TestCase_Name")){
			TestCase_Name_Index=i;
		}else if(cell.getStringCellValue().equalsIgnoreCase("User_ID")){
			User_ID_Index=i;
		}else if(cell.getStringCellValue().equalsIgnoreCase("Password")){
			Password_Index=i;
		}else if(cell.getStringCellValue().equalsIgnoreCase("BAN")){
			BAN_Index=i;
		}else if(cell.getStringCellValue().equalsIgnoreCase("Table_Data")){
			Table_Data_Index=i;
		}else if(cell.getStringCellValue().equalsIgnoreCase("Device")){
			Device_Index=i;
		}else if(cell.getStringCellValue().equalsIgnoreCase("Execute")){
			Execute_Index=i;
		}else if(cell.getStringCellValue().equalsIgnoreCase("Table_Data_Val")){
			Table_Data_Val_Index=i;
		}else if(cell.getStringCellValue().trim().equalsIgnoreCase("Result")){
			Result_Index=i;
			System.out.println(Result_Index);
		}else if(cell.getStringCellValue().trim().equalsIgnoreCase("Actual_Table_Data")){
			Actual_Table_Data_index=i;
			System.out.println(Actual_Table_Data_index);
		}else if(cell.getStringCellValue().trim().equalsIgnoreCase("Actual_Table_Data_Val")){
			Actual_Table_Data_Val_index=i;

		}i++;
		}
		
		
		//To get row number 
		
		int rowNumLen = ws.getLastRowNum() + 1;
		XSSFCell cellPassword,cellUser_ID,cell_Url,cellTestCase_Name,cellBAN,cellTable_Data,cellDevice,cell_Table_Data_Val,cell_Execute;
		
		int imgcount=0;
		try{
		  for (i = 1; i < rowNumLen; i++) {
			row = ws.getRow(i);
			/////////////     for converting to line 2 , i-1 -> i     //////////////			
			cell_Url = row.getCell(Url_Index);
			cellTestCase_Name = row.getCell(TestCase_Name_Index);
			cellUser_ID = row.getCell(User_ID_Index);
			cellPassword=row.getCell(Password_Index);
			cellBAN=row.getCell(BAN_Index);
			cellTable_Data = row.getCell(Table_Data_Index);
			cellDevice = row.getCell(Device_Index);
			cell_Table_Data_Val=row.getCell(Table_Data_Val_Index);
			cell_Execute=row.getCell(Execute_Index);
			String cellTestCase_Name1=cellTestCase_Name.toString();

			if (cell_Url !=null &&cell_Execute!=null&&cell_Execute.toString().equalsIgnoreCase("YES"))
			{
				
				//Creating Directory to Save Screenshots by TestCase Name
				File theDir = new File(outputPath);
				if (!theDir.exists()) {
					theDir.mkdir();     
				}
				File screenDir = new File(outputPath+"\\"+cellTestCase_Name1);
				if (!screenDir.exists()) {
					imgcount=0;screenDir.mkdir();     
				}
				
				image_Name = outputPath+"\\"+cellTestCase_Name1+"\\"+imgcount+".png";
				imgcount++;
				
				
				boolean browserExist=false;
				List<String> worklist=new ArrayList<String>();
				worklist.add("Open scanr url");
				worklist.add("Login with user role");
				worklist.add("Enter Ban");
				worklist.add("Open network_map");
				worklist.add("Click Device");
				worklist.add("Check Details");
				worklist.add("Close Scanr");
				
					for(String as:worklist)
					{		
						if (as.toLowerCase().contains("Open".toLowerCase()) && as.toLowerCase().contains("scanr".toLowerCase()) && as.toLowerCase().contains("url".toLowerCase())&& driver==null)
						{
							String workFlowLink = cell_Url.toString();
							driver=openScanr(workFlowLink);
						}
						else if (as.toLowerCase().contains("Login".toLowerCase()) && as.toLowerCase().contains("with".toLowerCase()) && as.toLowerCase().contains("User".toLowerCase()) && as.toLowerCase().contains("Role".toLowerCase())&&!browserExist)
						{
							browserExist=true;
							String login_id = cellUser_ID.toString();
							String login_password = cellPassword.toString();
							Login(login_id, login_password);
						}
						else if(as.toLowerCase().contains("enter".toLowerCase()) && as.toLowerCase().contains("Ban".toLowerCase()))
						{
							
							 Thread.sleep(2000);
							 DataFormatter df = new DataFormatter();
							//cellBAN.getRichStringCellValue();
							String Data = df.formatCellValue(cellBAN);
							 
							 
							 //Number n = null; 
							  enterBan(Data); 				
						}
						else if (as.toLowerCase().contains("open".toLowerCase()) && as.toLowerCase().contains("network_map".toLowerCase()))
						{
							network_map_Tab();
						}
						else if (as.toLowerCase().contains("click".toLowerCase()) && as.toLowerCase().contains("Device".toLowerCase()))
						{
							String Device_Type = cellDevice.toString();
							Click_Device(Device_Type);
						}
						else if(as.toLowerCase().contains("check".toLowerCase()) && as.toLowerCase().contains("Details".toLowerCase()))
						{
							 DataFormatter df = new DataFormatter();
							String Table_Data = df.formatCellValue(cellTable_Data);
							String Table_Data_Val = df.formatCellValue(cell_Table_Data_Val);
							Check_Details(Table_Data,Table_Data_Val,ExcelPath,row,wb);
							Snapshot(image_Name);
						}
						else if(as.toLowerCase().contains("close".toLowerCase()) && as.toLowerCase().contains("Scanr".toLowerCase()))
						{
							Close_Scanr();
						}

					}	
				
			}
			else if (cell_Url !=null && cell_Execute==null)
			{
				//		reportingTestName.setCellValue(cellTestName.toString());
				//		reportingStatus.setCellValue("No Run");	
				continue;
			}
			else if ((cell_Url.equals("") || cell_Url==null) && cell_Execute==null)
			{
				
				if(driver!=null)
				{
					
					driver.quit();
				}
				FileOutputStream fout = new FileOutputStream(excel);	
				wb.write(fout);
				fout.close();
				break;
			}
				
			}
		  	FileOutputStream fout = new FileOutputStream(excel);	
			wb.write(fout);
			fout.close();
			//break;
		  }// End of Try
			catch(Exception exception)
			{
				exception.printStackTrace();
				driver.close();
				driver.quit();
			}
		}

	public void Close_Scanr() throws Exception {
		Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
		driver.quit();
		driver=null;}
	
	
	public WebDriver openScanr(String url) throws Exception {
		System.setProperty("webdriver.ie.driver","D:\\E Drive Data\\For selenium\\IEDriverServer.exe");
		DesiredCapabilities caps = DesiredCapabilities.internetExplorer();
		caps.setCapability("ignoreZoomSetting", true);
		WebDriver driver = new InternetExplorerDriver(caps);
		// driver = new InternetExplorerDriver();
		Robot ignoreZoom = new Robot();
		ignoreZoom.keyPress(KeyEvent.VK_CONTROL);
		ignoreZoom.keyPress(KeyEvent.VK_0);
		ignoreZoom.keyRelease(KeyEvent.VK_CONTROL);
		driver.get(url);
		driver.getTitle();
		wait = new WebDriverWait(driver, 120000);
		return driver;
	}

	public void Login(String userName, String password) throws InterruptedException {
		try{
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/input")));
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/input"))	.clear();
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/input")).sendKeys(userName);
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/input")).clear();
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/input")).sendKeys(password);
		driver.findElement(By.xpath("/html/body/form[1]/center/table/tbody/tr[2]/td/center/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td/input")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"srv_successok\"]/input")));
		driver.findElement(By.xpath("//*[@id=\"srv_successok\"]/input")).click();	
		Thread.sleep(2000);}catch(Exception e){
			e.printStackTrace();
			driver.close();
			driver.quit();
		}
	}


	public void enterBan(String data) throws InterruptedException{
		try{
		ChangeTab();		
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("subscriberId")));
		driver.manage().window().maximize();
		driver.findElement(By.id("subscriberId")).clear();
		driver.findElement(By.id("subscriberId")).sendKeys(data);
		driver.findElement(By.xpath(("//input[contains(@src,'/isaac/web/images/cti/button_search.gif')]"))).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("banTabIFrame_0")));
		driver.switchTo().frame("banTabIFrame_0");}catch(Exception e){
			e.printStackTrace();
			driver.close();
			driver.quit();
		}
	}

	public void ChangeTab() throws InterruptedException{             
		try{
		Thread.sleep(3000);	
		String subWindowHandler = null;
		Set <String>handles = driver.getWindowHandles();
		Iterator<String> iterator = handles.iterator();
		while (iterator.hasNext()){
			subWindowHandler = iterator.next();
			driver.switchTo().window(subWindowHandler);
		}}catch(Exception e){
			e.printStackTrace();
			driver.close();
			driver.quit();
		}
	}
	
	public void Snapshot(String fileName){
		try {
			Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
			BufferedImage capture = new Robot().createScreenCapture(screenRect);
			ImageIO.write(capture, "png", new File(fileName));
			
		} 
		catch (IOException ex) {
			System.out.println(ex);
		} 
		catch (AWTException ex) {
			System.out.println(ex);
		}
		
	}
	

	public void SwitchFrame(String frameName) throws InterruptedException {
		try{
		System.out.println("New Frame");
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(frameName)));
		Thread.sleep(10000);
		driver.switchTo().frame(driver.findElement(By.xpath(frameName)));}catch(Exception e){
			e.printStackTrace();
			driver.close();
			driver.quit();
		}
	}

	// Function to open Network Map
	public void network_map_Tab() throws InterruptedException{
		try{
		wait.until(ExpectedConditions.elementToBeClickable(By.id("networkMapTab")));
		driver.findElement(By.id("networkMapTab")).click();
		driver.findElement(By.id("networkMapTab")).click();	
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='nwm_loadingMessage']/table/tbody/tr/td")));
		Thread.sleep(20000);
		System.out.println("Network Map Tab Visible");}catch(Exception e){
			e.printStackTrace();
			driver.close();
			driver.quit();
		}
	}
	
	
	public void Click_Device(String Device_Type) throws InterruptedException{
		try{
		//Clicking on Device
		SwitchFrame("//*[@id=\"nwm_contentFrame\"]");
		if(Device_Type.equalsIgnoreCase("RG"))																						//--------------------1.RG
							{
							wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='router0_status_wait']")));
							Thread.sleep(20000);
							
							/*driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/2wire5268_med.png']")).click();
							driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/2wire5268_med.png']")).click();
							driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/2wire5268_med.png']")).click();*/
						
							
							driver.findElement(By.xpath("//div[@id='mapholder']/div[@id='router0']/img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//div[@id='mapholder']/div[@id='router0']/img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//div[@id='mapholder']/div[@id='router0']/img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//img[@id='router0_image']")).click();
							driver.findElement(By.xpath("//img[@id='router0_image']")).click();
							
							}
		
		else if(Device_Type.equalsIgnoreCase("NM55"))
		{
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='router0_status_wait']")));
			driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/motorola_nm55_med.png']")).click();
			driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/motorola_nm55_med.png']")).click();
			driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/motorola_nm55_med.png']")).click();
			driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/motorola_nm55_med.png']")).click();

		}
		
		else if(Device_Type.equalsIgnoreCase("Ex_iNID"))																			//--------------------2.Ethernet
		{
		
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='nid0_status_wait']")));
		Thread.sleep(20000);
		driver.findElement(By.xpath("//img[@id='nid0_image']")).click();
		driver.findElement(By.xpath("//img[@id='nid0_image']")).click();
		}
		
		else if(Device_Type.equalsIgnoreCase("iNID"))																			//--------------------2.Ethernet
		{
			
		Thread.sleep(20000);
		driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/i38_med.png']")).click();
		driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/i38_med.png']")).click();
		}
		
		else if(Device_Type.equalsIgnoreCase("Ethernet"))																			//--------------------2.Ethernet
							{
							
							wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/cisco_ven501_med.png']")));
							Thread.sleep(20000);
							driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/cisco_ven501_med.png']")).click();
							driver.findElement(By.xpath("//img[@src='/isaac/web/images/cti/HTMLMap/med/cisco_ven501_med.png']")).click();
							}
		else if(Device_Type.equalsIgnoreCase("Ruckus"))																			//--------------------2.Ruckus
							{
							wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[contains(@id,'status_wait']")));
							driver.findElement(By.xpath("//*[contains(@src,'ruckus_sm')]")).click();}
		else if(Device_Type.equalsIgnoreCase("Wireless Access Point"))																//--------------------3.WAP
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[contains(@id,'status_wait']")));
							driver.findElement(By.xpath("//*[contains(@id,'Wireless_AP')]")).click();}
		else if(Device_Type.equalsIgnoreCase("Att_Mobility_Equipment"))																//--------------------4.Mobility Equipment
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[contains(@id,'status_wait']")));
							driver.findElement(By.xpath("//*[contains(@src,'group_mobility')]")).click();
							 driver.findElement(By.xpath("//*[contains(@src,'new_cellphone_med')][1]")).click();}
		else if(Device_Type.equalsIgnoreCase("NID"))																				//--------------------5.NID
							{
							wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='nid0_status_wait']")));
							driver.findElement(By.xpath("//*[@id='nid0_image']")).click();}	
		else if(Device_Type.equalsIgnoreCase("DSLAM"))																				//--------------------6.DSLAM
							{
							wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='vrad0_status_wait']")));
							driver.findElement(By.xpath("//*[@id='vrad0_image']")).click();
							driver.findElement(By.xpath("//*[@id='vrad0_image']")).click();}	
		else if(Device_Type.equalsIgnoreCase("Non_Att_Wireless_Device"))															//--------------------7.Non Att Wireless Device
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[contains(@id,'status_wait']")));
							driver.findElement(By.xpath("//*[contains(@src,'nonATT')]")).click();
							 driver.findElement(By.xpath("//*[contains(@src,'pcWifi_med')][1]")).click();}
		else if(Device_Type.equalsIgnoreCase("Non_Att_Ethernet_Device"))															//--------------------8.Non Att Ethernet Device
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[contains(@id,'status_wait']")));
							driver.findElement(By.xpath("//*[contains(@src,'nonATT')]")).click();
							 driver.findElement(By.xpath("//*[contains(@src,'pc_med')][1]")).click();}
		else if(Device_Type.equalsIgnoreCase("Uverse_Voice_Line"))																	//--------------------9.Uverse Voice Line
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[contains(@id,'status_wait']")));
							driver.findElement(By.xpath("//*[contains(@src,'uverseVoice')]")).click();
							 driver.findElement(By.xpath("//*[@id='phone0_image']")).click();}	
		else if(Device_Type.equalsIgnoreCase("STB"))																				//--------------------10.STB
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[contains(@id,'status_wait']")));
							driver.findElement(By.xpath("//*[contains(@src,'uverseReceiver')]")).click();
							 driver.findElement(By.xpath("//*[contains(@src,'stb')][1]")).click();}
		else if(Device_Type.equalsIgnoreCase("OWA"))																				//--------------------11.OWA
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='owa0_status_wait']")));
							Thread.sleep(20000);
							driver.findElement(By.xpath("//*[@id='owa0_image']")).click();
							driver.findElement(By.xpath("//*[@id='owa0_image']")).click();}
		else if(Device_Type.equalsIgnoreCase("Cell_Site"))																				//--------------------11.Cell Site
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='ctw0_status_wait']")));
							Thread.sleep(20000);
							driver.findElement(By.xpath("//*[@id='ctw0_image']")).click();
							driver.findElement(By.xpath("//*[@id='ctw0_image']")).click();}
		else if(Device_Type.equalsIgnoreCase("POE"))																					//--------------------11.POE
							{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='poe0_status_wait']")));
							Thread.sleep(20000);
							driver.findElement(By.xpath("//*[@id='poe0_image']")).click();
							driver.findElement(By.xpath("//*[@id='poe0_image']")).click();
							
							}
		
		else if(Device_Type.equalsIgnoreCase("Voice_Lines"))																					//--------------------11.POE
		{wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='poe0_status_wait']")));
		Thread.sleep(20000);
		driver.findElement(By.xpath("//*[@id='generic_voice_image']")).click();
		
		driver.findElement(By.xpath("//*[@id='phone1_image']")).click();
		
		}
		
		
		
		}catch(Exception e){
			e.printStackTrace();
			driver.close();
			driver.quit();
		}}

   // Checking device details code
	
	@SuppressWarnings("unused")
	public void Check_Details(String Table_Data,String Table_Data_Val,String ExcelPath,XSSFRow rowi,XSSFWorkbook wb) throws InterruptedException, IOException{
		
	   try{	
	   String[] Tokens = Table_Data.split(",");
	   String[] Token_Val = Table_Data_Val.split(",");
	   
	   //File excel = new File(ExcelPath);
	   XSSFCell cell_Result,cellActual_Table_Data,cellActual_Table_Data_Val;	
	   cell_Result=rowi.createCell(Result_Index);
	   
	   cellActual_Table_Data=rowi.createCell(Actual_Table_Data_index);
	   cellActual_Table_Data_Val=rowi.createCell(Actual_Table_Data_Val_index);
	   	   
	   //Creating a hashMap for expected device details 
	   HashMap<String, String> Expected_ColumnValueMap = new HashMap<String, String>();
	   int flag=0;
	   System.out.println("Expected_Map");
	   for(String Ex_token:Tokens){
		   if(Token_Val[flag].equals(""))
			   Expected_ColumnValueMap.put(Ex_token.toString().trim(), null) ;
		   else{
			   
			   Expected_ColumnValueMap.put(Ex_token.toString().trim(), Token_Val[flag].toString().trim()) ;}
		   flag++;
		   System.out.println(Ex_token+"  "+Expected_ColumnValueMap.get(Ex_token)+"\n");
	   }
 	   
	   //Checking of Actual results
	    ChangeTab();
	    SwitchFrame("//*[@id=\"banTabIFrame_0\"]");
		SwitchFrame("//*[@id=\"nwm_detailFrame\"]");
		
	    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='device_details']")));
	  	WebElement Table = driver.findElement(By.xpath("//*[@id='device_details']"));
		
		 
		 List<WebElement> tableRows=Table.findElements(By.tagName("tr"));
		 
		 List<String> rowColumnsList=new ArrayList<String>();
		 HashMap<String, String> columnValueMap=new HashMap<String, String>();
		 
		 
		 System.out.println("Actual Column Value Map");
		 
		 int count = tableRows.size();
		 
		 Date date = new Date();
		 SimpleDateFormat sdf = new SimpleDateFormat("MMddyyyy_hmmss_a");
		 String formattedDate;
		 System.out.println(sdf.format(date));
		 WebElement E;

		 
		 for(int j=0;j<Token_Val.length;j++){
		 
			 for(int i =0; i<count; i++) 
			 {	E=tableRows.get(i);
				 if (E.getText().contains(Token_Val[j])){
				 	 tableRows.get(i).click();	
				 	 Snapshot(sdf.format(date));}}
		 }
			 
		 int Token_Val_Length = Token_Val.length-1;
		 int Tokens_Length = Tokens.length-1;
		 
		  //Creating HashMap for actual Results Table
		 if(count>0  && Token_Val_Length+1>0 && Tokens_Length+1 >0){			 
			 for(WebElement row:tableRows)
			 {		 
				 rowColumnsList.add(row.getText());
				 
			 }
		 
			 for(String rowText: rowColumnsList )
			 {
				 String text[]=rowText.split(": ");
				 if(text.length==1 && !(text.equals("")) )
				 {
					 if (columnValueMap.containsKey(text[0].toString().trim()))
					 {
						 duplicate = duplicate+1;
						 columnValueMap.put((text[0].toString().trim()+"_"+duplicate).toString(),null); 
					 }
					 else{
						 columnValueMap.put(text[0].toString().trim(),null);
					
					 }
				 }
			 else if(text.length==2)
			 {
				 if (columnValueMap.containsKey(text[0].toString().trim()))
				 {
					 duplicate = duplicate+1;
					 columnValueMap.put((text[0].toString().trim()+"_"+duplicate).toString(),text[1].toString().trim()); 
				 }
				 else{
				 columnValueMap.put(text[0].toString().trim(),text[1].toString().trim());
				 }
			 }
			 System.out.println(text[0].toString().trim()+"  "+columnValueMap.get(text[0].toString().trim())+"\n");
			 
		 }}
		 else if(count == 0 && Token_Val_Length==0 && Tokens_Length ==0){
			 cell_Result.setCellValue("PASS : Table Rows null for the Image");
			 
		 }else if(count>0 && Token_Val_Length==0 && Tokens_Length ==0){
			 cell_Result.setCellValue("PASS : Only Image Existence check Required");}
		 
		 else
			 cell_Result.setCellValue("FAIL : Expected Value null");
			 //End of Create Actual Results HashMap Table

		
			 
		 //Creating a hashMap for entering searched i.e Test Results 
		  HashMap<String, String> Result_ColumnValueMap = new HashMap<String, String>();
		  
		 
		 //Searching of Keys of Expected Results in Actual Results
		 for(String Ex_keys: Expected_ColumnValueMap.keySet())
		 {
			 
			 if(columnValueMap.containsKey(Ex_keys))
			 {
				 
				  if(Expected_ColumnValueMap.get(Ex_keys).equalsIgnoreCase(columnValueMap.get(Ex_keys)))
				 {
					 Result_ColumnValueMap.put(Ex_keys, columnValueMap.get(Ex_keys));
				 } else
					 Result_ColumnValueMap.put(Ex_keys, "Value not found");
			 }
			  
		 }//End of Search Loop
		
		 for(String Result: Result_ColumnValueMap.keySet())
			 System.out.println(Result+"   :    "+Result_ColumnValueMap.get(Result));
		 
		// ________________________________________________________________________________________________________________________________________________________
		 	// Write Result on Excel
		 					
				String token_list="";
				String tokenval_list="";
				
				int flag_comma=0;
				for (String token:Result_ColumnValueMap.keySet() ) {
					if(flag_comma==0){token_list = token; flag_comma++; tokenval_list =Result_ColumnValueMap.get(token); continue;}
					token_list=token_list+","+token;
					tokenval_list=tokenval_list+","+Result_ColumnValueMap.get(token);
				}//end of for loop
				
				/*int flag_click=1;	
				 for(String rowText: rowColumnsList )
				 {
					 String[] rowText_split=token_list.split(":");
					 String rt_s=rowText_split[1];
				for (String rt:Result_ColumnValueMap.keySet() ) {
					String[] Tokens_=token_list.split(",");
					if(Tokens_[flag_click].equalsIgnoreCase(rt_s))
						{flag_click++;}
					}
				 }*/
				 
				 
				 
				cellActual_Table_Data_Val.setCellValue(tokenval_list);
				cellActual_Table_Data.setCellValue(token_list);
				cellActual_Table_Data_Val.setCellValue(tokenval_list);
				
				if(Result_ColumnValueMap.containsValue("Value not found") || Result_ColumnValueMap.isEmpty())
				 	cell_Result.setCellValue("FAIL");
				else
					cell_Result.setCellValue("PASS");
					
					
					
		}catch(Exception e){
			e.printStackTrace();
			driver.close();
			driver.quit();
		}
	}// End of Check Details Function
	

public static void main(String[] args) throws IOException, InterruptedException{
	Scanr_NM auto = new Scanr_NM();
	auto.Read("D:\\E Drive Data\\For selenium\\1607_TestCasesAuto_SCANR.xlsx");
}

}