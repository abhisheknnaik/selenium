package com.mastercardOffers.main;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.ProxySelector;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.Proxy.ProxyType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;
import ru.yandex.qatools.ashot.screentaker.ViewportPastingStrategy;

import com.google.common.io.Files;

public class Input {
	public static WebDriver driver = null;
	static int offerId = 0;
	static String outputFolder;

	@SuppressWarnings({ "resource" })
	public static void main(String[] args) throws Exception {
		 chrome();
		//ie();
		//edgeBrowser();
		String str = getUniqueNumber();

		outputFolder = "src\\main\\resources\\output\\" + str;
		new File(outputFolder).mkdir();

		// driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
		XSSFWorkbook workbook = new XSSFWorkbook();
		String inputXls = "src\\main\\resources\\daily_offercheck.xlsx";

		// read 'offer_page' sheet
		LinkedHashMap<String, Object[]> map = readXlsxFile(inputXls,
				"offer_page");
		map = generateOutputForOffers(map);
		workbook = createSheet(map, "offer_page", workbook);

		// read 'search' sheet
		LinkedHashMap<String, Object[]> map1 = readXlsxFile(inputXls, "search");
		map1 = generateOutputForsearch(map1);
		workbook = createSheet(map1, "search", workbook);

		System.out.println(str);
		// create work book
		createWorkBook(outputFolder + "\\daily_offercheck_out" + str + ".xlsx",
				workbook);
	}

	public static void chrome() throws URISyntaxException {

		ChromeOptions chromeOptions = new ChromeOptions();

		String userhome = "an00542992";// Advapi32Util.getUserName(); //
		System.getProperty("user.home");
		System.setProperty("webdriver.chrome.driver",
				"src\\main\\resources\\chromedriver.exe");
		chromeOptions.setBinary("C:\\Users\\" + userhome
				+ "\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe");
		driver = new ChromeDriver(chromeOptions);
	}

	private static java.net.Proxy findProxy(URI uri) {
		try {
			ProxySelector selector = ProxySelector.getDefault();
			List<java.net.Proxy> proxyList = selector.select(uri);
			if (proxyList.size() > 1)
				return proxyList.get(0);

		} catch (IllegalArgumentException e) {
		}
		return java.net.Proxy.NO_PROXY;

	}

	public static void ie() throws URISyntaxException {

		URI url = new URI("http://tmgate.techm/wpad.dat");

		java.net.Proxy p = findProxy(url);
		System.setProperty("webdriver.ie.driver",
				"src\\main\\resources\\\\IEDriverServer.exe");
		DesiredCapabilities cap = DesiredCapabilities.internetExplorer();

		cap.setCapability(
				InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,
				true);
		cap.setCapability("ignoreProtectedModeSettings", true);
		cap.setCapability("initialBrowserUrl", "www.google.co.in");
		cap.setCapability("InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION",
				true);
		cap.setCapability("ignoreZoomSetting", true);
		// cap.setProxy(p);
		driver = new InternetExplorerDriver(cap);

	}

	public static void edgeBrowser() throws URISyntaxException {

		System.setProperty("webdriver.edge.driver",
				"src\\main\\resources\\\\MicrosoftWebDriver14_14392.exe");

		driver = new EdgeDriver();

	}

	public static String getUniqueNumber() {
		Date date = new Date();
		// *** same for the format String below
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy_HH-mm");
		String str = dateFormat.format(date);
		return str;
	}

	public static void clickElement(WebElement element)
			throws InterruptedException {
		Thread.sleep(5000);
		WebDriverWait wait = new WebDriverWait(driver, 60);
		wait.until(ExpectedConditions.elementToBeClickable(element));
		JavascriptExecutor js = ((JavascriptExecutor) driver);
		js.executeScript("arguments[0].scrollIntoView();", element);
		Thread.sleep(500);
		js.executeScript("window.scrollBy(0, -200);");

		Thread.sleep(500);
		new Actions(driver).moveToElement(element).perform();

		Thread.sleep(500);
		element.click();
		Thread.sleep(500);

	}

	public static LinkedHashMap<String, Object[]> generateOutputForOffers(
			LinkedHashMap<String, Object[]> map) throws Exception {
		LinkedHashMap<String, Object[]> offerMap = new LinkedHashMap<String, Object[]>();
		// loop a Map
		for (Map.Entry<String, Object[]> entry : map.entrySet()) {
			Object[] lst = entry.getValue();
			String key = entry.getKey();
			System.out
					.println("Key : " + entry.getKey() + " Value : " + lst[1]);

			if (lst[1].toString().equalsIgnoreCase("Page url")) {
				offerMap.put(key, lst);
			} else {
				String status = verifyOfferUrl(lst[1].toString());
				offerMap.put(key,
						new Object[] { key, lst[1].toString(), status });
			}
		}
		return offerMap;
	}

	public static LinkedHashMap<String, Object[]> generateOutputForsearch(
			LinkedHashMap<String, Object[]> map) throws Exception {
		LinkedHashMap<String, Object[]> offerMap = new LinkedHashMap<String, Object[]>();
		// loop a Map
		for (Map.Entry<String, Object[]> entry : map.entrySet()) {
			Object[] lst = entry.getValue();
			String key = entry.getKey();
			System.out
					.println("Key : " + entry.getKey() + " Value : " + lst[1]);

			if (lst[1].toString().equalsIgnoreCase("Page url")) {
				offerMap.put(key, lst);
			} else {
				Object[] status = verifySearchUrl(lst);
				offerMap.put(key, status);
			}
		}
		return offerMap;
	}

	static LinkedHashMap<String, Object[]> readXlsxFile(String fileName,
			String SheetName) throws IOException {
		String key = null;
		Object value = null;
		InputStream ExcelFileToRead = new FileInputStream(fileName);
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
		LinkedHashMap<String, Object[]> linkedHashMap = new LinkedHashMap<String, Object[]>();

		XSSFSheet sheet = wb.getSheet(SheetName);
		XSSFRow row;
		XSSFCell cell;

		Iterator<Row> rows = sheet.rowIterator();
		List<Object> list = new ArrayList<Object>();
		while (rows.hasNext()) {
			row = (XSSFRow) rows.next();
			Iterator<Cell> cells = row.cellIterator();
			while (cells.hasNext()) {

				cell = (XSSFCell) cells.next();
				if (cell.getColumnIndex() == 0) {
					key = getXssFCellValue(cell).trim();
					list.add(getXssFCellValue(cell).trim());
				}

				else {
					if (getXssFCellValue(cell) != null) {
						list.add(getXssFCellValue(cell).trim());
					}
				}
			}
			if (list.size() > 0) {
				value = list.toArray();
				list.clear();
				linkedHashMap.put(key, (Object[]) value);
			}

		}
		System.out.println("size " + linkedHashMap.size());
		System.out.println(linkedHashMap);
		wb.close();
		return linkedHashMap;
	}

	static String getXssFCellValue(XSSFCell cell) {
		if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
			return cell.getStringCellValue();
		} else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
			return Double.toString(cell.getNumericCellValue());
		}
		return null;
	}

	static String verifyOfferUrl(String url) throws Exception {

		 driver.get(url);
	//	driver.navigate().to(url);
		driver.manage().window().maximize();
		Thread.sleep(1000);
		System.out.println(driver.getTitle());

		takeSnapShot(driver, outputFolder + "\\offer_" + offerId + ".png");
		offerId = offerId + 1;
		if (driver.getTitle().equalsIgnoreCase("Offer")) {
			return "Yes";
		} else {
			return "No";
		}
	}

	static Object[] verifySearchUrl(Object[] lst) throws Exception {

		String linkText, tabContentId, status = null, cssForNoResult, cssForError;
		driver.get(lst[1].toString());
		// driver.navigate().to(lst[1].toString());
		Object[] updatedListObjects = new Object[lst.length];

		driver.manage().window().maximize();
		String title = driver.getTitle();
		System.out.println(driver.getTitle());
		if (title.equalsIgnoreCase("This page can’t be displayed")) {
			Thread.sleep(2000);
			driver.get(lst[1].toString());
			title = driver.getTitle();
			System.out.println(driver.getTitle());
		}

		updatedListObjects[0] = lst[0];
		updatedListObjects[1] = lst[1];

		for (int i = 2; i < lst.length; i++) {
			System.out.println(lst[i].toString());
			linkText = lst[i].toString().trim();
			if (linkText.equalsIgnoreCase("na")) {
				status = "NA";
			} else {
				Thread.sleep(2000);
				clickElement(driver.findElement(By.linkText(linkText)));
				tabContentId = driver.findElement(By.linkText(linkText))
						.getAttribute("href");
				tabContentId = tabContentId.split("#")[1];
				cssForNoResult = "#" + tabContentId + " .hasNoResults";
				System.out.println(cssForNoResult);
				String screenshotFileName = null;
				status = driver.findElement(By.cssSelector(cssForNoResult))
						.getAttribute("style");

				// retry to get status
				if (status.trim().isEmpty()) {
					Thread.sleep(1000);
					status = driver.findElement(By.cssSelector(cssForNoResult))
							.getAttribute("style");
				}

				if (status.trim().isEmpty()) {
					cssForError = "#" + tabContentId
							+ " .serviceDownError span";
					String error = driver.findElement(
							By.cssSelector(cssForError)).getText();
					System.out.println("error " + error);

					if (!error.trim().isEmpty()) {
						status = "Error : " + error;
						screenshotFileName = "Error";
					}

				} else if (status.equalsIgnoreCase("display: block;")) {
					status = "No Results";
					screenshotFileName = "No Results";
				} else {
					status = "Available";
					screenshotFileName = "Available";
				}
				Thread.sleep(1000);
				takeSnapShot(driver, outputFolder + "\\search_" + title + "_"
						+ linkText + "_" + screenshotFileName + ".png");
			}
			System.out.println(status);
			updatedListObjects[i] = status;
		}
		return updatedListObjects;

	}

	static public XSSFWorkbook createSheet(LinkedHashMap<String, Object[]> map,
			String sheetName, XSSFWorkbook workBook) throws IOException {
		// Blank workbook
		XSSFWorkbook workbook = workBook;

		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet(sheetName);

		// This data needs to be written (Object[])
		Map<String, Object[]> data = map;
		String strCellValue;
		// Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		int rownum = 0, cellnum = 0;
		;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				CellStyle style = workbook.createCellStyle();
				style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				style.setBorderTop(HSSFCellStyle.BORDER_THIN);
				style.setBorderRight(HSSFCellStyle.BORDER_THIN);
				style.setBorderLeft(HSSFCellStyle.BORDER_THIN);

				if (obj instanceof String) {
					strCellValue = (String) obj.toString().trim();
					if (strCellValue.equalsIgnoreCase("No Results")
							|| strCellValue.equalsIgnoreCase("No")) {
						style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW
								.getIndex());
						style.setFillPattern(CellStyle.SOLID_FOREGROUND);
						cell.setCellStyle(style);
					}
					if (strCellValue.toLowerCase().contains("error")) {
						style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE
								.getIndex());
						style.setFillPattern(CellStyle.SOLID_FOREGROUND);
						cell.setCellStyle(style);
					}
					if (strCellValue.equalsIgnoreCase("NA")) {
						style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT
								.getIndex());
						style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					}

					cell.setCellStyle(style);
					cell.setCellValue((String) obj);
				} else if (obj instanceof Integer) {
					cell.setCellValue((Integer) obj);
				}
			}
		}
		for (int i = 1; i <= cellnum + 1; i++) {
			sheet.autoSizeColumn(i);
		}
		return workbook;
	}

	public static void createWorkBook(String fileName, XSSFWorkbook workbook) {
		try {
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File(fileName));
			workbook.write(out);
			out.close();
			System.out.println("out .xls file written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void takeSnapShot(WebDriver webdriver, String fileWithPath)
			throws Exception {

		// Convert web driver object to TakeScreenshot
		TakesScreenshot scrShot = ((TakesScreenshot) webdriver);

		// Call getScreenshotAs method to create image file
		File SrcFile = scrShot.getScreenshotAs(OutputType.FILE);

		// Move image file to new destination
		File DestFile = new File(fileWithPath);

		// Copy file at destination
		Files.copy(SrcFile, DestFile);

		// // take full screen shot for chrome
		// WebDriver driver=webdriver;
		// final Screenshot screenshot = new AShot().shootingStrategy(
		// new ViewportPastingStrategy(500)).takeScreenshot(driver);
		// final BufferedImage image = screenshot.getImage();
		// ImageIO.write(image, "PNG", new File(fileWithPath));

	}

}
