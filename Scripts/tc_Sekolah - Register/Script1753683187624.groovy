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
import custom.Verif as Verif


WebUI.openBrowser('')

WebUI.navigateToUrl('https://sekolah.atrialfa.shop/', FailureHandling.STOP_ON_FAILURE)

WebUI.maximizeWindow()

WebUI.click(findTestObject("Object Repository/Page_SMPN6 - Laravel/a_Register"))

String inputNama = "Abdul dua"

WebUI.setText(findTestObject("Object Repository/Page_Register - Laravel/input_Name_name"), inputNama)

WebUI.setText(findTestObject("Object Repository/Page_Register - Laravel/input_Email_email"), "abdul@dua.mantapdua")

WebUI.setText(findTestObject("Object Repository/Page_Register - Laravel/input_Password_password"), "muantap12345")

WebUI.setText(findTestObject("Object Repository/Page_Register - Laravel/input_Confirm Password_password_confirmation"), "muantap12345")

WebUI.click(findTestObject("Object Repository/Page_Register - Laravel/button_Register"))

String excelTC = "C:\\Users\\malik\\Documents\\Project Abdul\\Atry Katalon Project\\Excel\\TC - Training sekolah.xlsx"

Verif.checkLoginAndWriteExcelRegister(excelTC, inputNama)


