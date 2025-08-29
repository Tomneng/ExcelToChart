package backend.exceltochart.service

import backend.exceltochart.config.ApiResponse
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import java.io.BufferedInputStream
import java.io.File
import java.io.FileOutputStream


@Service
class ExcelParsingService {
    fun validateExcel(file: MultipartFile): ApiResponse {

        val excelFile = multipartToFile(file)
        val workbook = XSSFWorkbook(excelFile)
        val sheet = workbook.getSheetAt(0)
        val rowIterator: MutableIterator<Row?>? = sheet.iterator()
        while (rowIterator!!.hasNext()) {
            val row: Row = rowIterator.next()!!

            // 각각의 행에 존재하는 모든 열(cell)을 순회한다.
            val cellIterator = row.cellIterator()

            while (cellIterator.hasNext()) {
                val cell = cellIterator.next()

                // cell의 타입을 하고, 값을 가져온다.
                when (cell.getCellType()) {
                    CellType.NUMERIC -> print(
                        cell.getNumericCellValue().toInt().toString() + "\t"
                    ) //getNumericCellValue 메서드는 기본으로 double형 반환
                    CellType.STRING -> print(cell.getStringCellValue() + "\t")
                    CellType._NONE -> print("none")
                    CellType.FORMULA -> print("Formula")
                    CellType.BLANK -> print("blank")
                    CellType.BOOLEAN -> print("boolean")
                    CellType.ERROR -> print("Error")
                }
            }
            println()
        }
        return ApiResponse(true,"성공","없음")
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