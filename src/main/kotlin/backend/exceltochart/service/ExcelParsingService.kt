package backend.exceltochart.service

import backend.exceltochart.config.ApiResponse
import org.apache.commons.lang3.StringUtils.center
import org.apache.poi.ss.formula.SheetRange
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import java.io.BufferedInputStream
import java.io.File
import java.io.FileOutputStream


@Service
class ExcelParsingService(
    private val excelAnalyzeService: ExcelAnalyzeService
) {

    fun validateExcel(file: MultipartFile): ApiResponse {

        val excelFile = multipartToFile(file)
        val workbook = XSSFWorkbook(excelFile)
        val sheet = workbook.getSheetAt(0)
        return ApiResponse(true,"성공","없음")
    }

    fun processParsing(file: MultipartFile): ApiResponse {
        val excelFile = multipartToFile(file)
        val workbook = XSSFWorkbook(excelFile)
        val sheet = workbook.getSheetAt(0)
        val center = detectSheetRange(sheet)
        print("${center[0]} ${center[1]}")
        excelAnalyzeService.nineCellExtractSafe(sheet, center)

        return ApiResponse(true,"성공","없음")
    }

    fun detectSheetRange(sheet: Sheet): Array<Int> {
        val lastRow = sheet.getRow(sheet.lastRowNum)
        val lastColumnNum = lastRow.lastCellNum
        val rangeArray = Array(2, init = { 0 })
        rangeArray[0] = sheet.lastRowNum
        rangeArray[1] = lastColumnNum.toInt()
        return rangeArray
    }


    fun multipartToFile(multipartFile: MultipartFile): File {
        val tempFile = File.createTempFile("temp", multipartFile.originalFilename?.let { "-$it" })

        // MultipartFile의 InputStream을 이용해 File 객체를 생성합니다.
        BufferedInputStream(multipartFile.inputStream).use { inputStream ->
            FileOutputStream(tempFile).use { outputStream ->
                inputStream.copyTo(outputStream)
            }
        }
        return tempFile
    }


}