package custom

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows

import internal.GlobalVariable

import java.sql.*
import org.apache.poi.xssf.usermodel.*
import java.io.FileOutputStream
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell
import java.text.SimpleDateFormat
import java.util.Date

public class Verif {

	private static CellStyle createCommonBorderedStyle (Workbook workbook) {
		CellStyle style = workbook.createCellStyle()
		style.setBorderTop(BorderStyle.THIN)
		style.setBorderBottom(BorderStyle.THIN)
		style.setBorderLeft(BorderStyle.THIN)
		style.setBorderRight(BorderStyle.THIN)
		style.setAlignment(HorizontalAlignment.LEFT)
		style.setVerticalAlignment(VerticalAlignment.TOP)
		style.setWrapText(true)
		return style
	}

	@Keyword
	def static void checkLoginAndWriteExcel(String excelPath) {
		boolean verif = false

		while (verif==false) {
			verif = WebUI.verifyTextPresent("You're logged in!", false, FailureHandling.STOP_ON_FAILURE)
		}

		String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date())
		String fileCapturePath = "C:\\Users\\malik\\Documents\\Project Abdul\\Atry Katalon Project\\Capture\\Capture_Positive_${timestamp}.png"

		if (verif) {
			WebUI.takeScreenshot(fileCapturePath, FailureHandling.STOP_ON_FAILURE)
			FileInputStream file = new FileInputStream(excelPath)
			Workbook workbook = new XSSFWorkbook(file)
			CellStyle borderedStyle = createCommonBorderedStyle(workbook)
			Sheet sheet = workbook.getSheetAt(0)

			Row row = sheet.getRow(9)
			Cell cell = row.createCell(7)
			cell.setCellValue("As Expected")
			cell.setCellStyle(borderedStyle)

			file.close()
			FileOutputStream outFile = new FileOutputStream(excelPath)
			workbook.write(outFile)
			outFile.close()
		}
	}
	@Keyword
	def static void checkLoginAndWriteExcelNegative(String excelPath) {
		boolean verif = false

		while (verif==false) {
			verif = WebUI.verifyTextPresent("These credentials do not match our records.", false, FailureHandling.STOP_ON_FAILURE)
		}

		String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date())
		String fileCapturePath = "C:\\Users\\malik\\Documents\\Project Abdul\\Atry Katalon Project\\Capture\\Capture_Negative_${timestamp}.png"

		if (verif) {
			WebUI.takeScreenshot(fileCapturePath, FailureHandling.STOP_ON_FAILURE)
			FileInputStream file = new FileInputStream(excelPath)
			Workbook workbook = new XSSFWorkbook(file)
			Sheet sheet = workbook.getSheetAt(0)

			CellStyle borderedStyle = createCommonBorderedStyle(workbook)

			Row row = sheet.getRow(10)
			Cell cell = row.createCell(7)
			cell.setCellValue("As Expected")
			cell.setCellStyle(borderedStyle)

			file.close()
			FileOutputStream outFile = new FileOutputStream(excelPath)
			workbook.write(outFile)
			outFile.close()
		}
	}
	@Keyword
	def static void checkLoginAndWriteExcelRegister(String excelPath, String inputNama) {
		String url = GlobalVariable.DB_URL
		String username = GlobalVariable.DB_Username
		String password = GlobalVariable.DB_Pass

		Class.forName("com.mysql.cj.jdbc.Driver")//tulis aja

		Connection conn = DriverManager.getConnection(url, username, password)
		Statement stmt = conn.createStatement()
		ResultSet rs1 = stmt.executeQuery("select name from sekolah.users order by id desc limit 10")


		boolean verif = false

		while (verif==false) {
			verif = WebUI.verifyTextPresent("You're logged in!", false, FailureHandling.STOP_ON_FAILURE)
		}

		String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date())
		String fileCapturePath = "C:\\Users\\malik\\Documents\\Project Abdul\\Atry Katalon Project\\Capture\\Capture_Register_${timestamp}.png"

		if (verif) {
			WebUI.takeScreenshot(fileCapturePath, FailureHandling.STOP_ON_FAILURE)
			FileInputStream file = new FileInputStream(excelPath)
			Workbook workbook = new XSSFWorkbook(file)
			Sheet sheet = workbook.getSheetAt(0)
			CellStyle borderedStyle = createCommonBorderedStyle(workbook)


			Row row = sheet.getRow(11)
			Cell cell = row.createCell(7)
			cell.setCellValue("As Expected")
			cell.setCellStyle(borderedStyle)
			Cell cell2 = row.createCell(15)

			boolean isNameFoundDB = false
			while(rs1.next()) {
				if(rs1.getString("name").equals(inputNama)) {
					isNameFoundDB = true
					break
				}
			}
			if (isNameFoundDB) {
				cell2.setCellValue("Sudah di Verifikasi di database : Register berhasil masuk ke Database (ditemukan dalam 10 data terbaru)")
				cell2.setCellStyle(borderedStyle)
			}else {
				cell2.setCellValue("'Verifikasi database GAGAL: Nama TIDAK ditemukan dalam 10 data terbaru'")
			}

			file.close()
			FileOutputStream outFile = new FileOutputStream(excelPath)
			workbook.write(outFile)
			outFile.close()
		}
	}
}
