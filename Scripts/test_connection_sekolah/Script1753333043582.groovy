import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys

import java.sql.*
import org.apache.poi.xssf.usermodel.*
import java.io.FileOutputStream
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell
import java.text.SimpleDateFormat
import java.sql.Date
import java.util.Date


String url = GlobalVariable.DB_URL 
String username = GlobalVariable.DB_Username                                
String password = GlobalVariable.DB_Pass           

// Optional (boleh ditulis, boleh enggak):
Class.forName("com.mysql.cj.jdbc.Driver")

Connection conn = DriverManager.getConnection(url, username, password)

Statement stmt = conn.createStatement()

stmt.execute("USE sekolah")
ResultSet rs1 = stmt.executeQuery("select * from sekolah.users order by id desc limit 10")
ResultSetMetaData meta=rs1.getMetaData()
int columnCount = meta.getColumnCount()

XSSFWorkbook workbook = new XSSFWorkbook()
XSSFSheet sheet = workbook.createSheet("Data Users")


XSSFRow headerRow = sheet.createRow(0)
for (int ii=1;ii<=columnCount;ii++) {
	headerRow.createCell(ii-1).setCellValue(meta.getColumnName(ii))
}

int rowIndex = 1
while (rs1.next()) {
	XSSFRow row = sheet.createRow(rowIndex++)
	for (int ii=1;ii<=columnCount;ii++) {
		def value = rs1.getString(ii)
		row.createCell(ii-1).setCellValue(value)
	}
}

String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date())

FileOutputStream fileOut = new FileOutputStream("C:/Users/malik/Documents/Project Abdul/Atry Katalon Project/Excel/Data Users_${timestamp}.xlsx")
workbook.write(fileOut)

conn.close()
