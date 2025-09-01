package backend.exceltochart.service

import backend.exceltochart.config.ApiResponse
import backend.exceltochart.model.CellRole
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import java.io.BufferedInputStream
import java.io.File
import java.io.FileOutputStream
import kotlin.collections.flatten


@Service
class ExcelParsingService(
    private val excelAnalyzeService: ExcelAnalyzeService,
    private val cellInfoUtilService: CellInfoUtilService,
    private val cellClassificationService: CellClassificationService
) {

    fun processParsing(file: MultipartFile): ApiResponse {
        val excelFile = multipartToFile(file)
        val workbook = XSSFWorkbook(excelFile)
        val sheet = workbook.getSheetAt(0)
        val fullRange = cellInfoUtilService.detectSheetRange(sheet)

        // 전체 시트 분류 실행
        val roleMatrix = classifyFullSheet(sheet, fullRange)

        // 결과 출력
        cellInfoUtilService.printRoleMatrix(roleMatrix)

        // 추가 통계 정보
        val headerCount = roleMatrix.flatten().count { it == CellRole.HEADER }
        val dataCount = roleMatrix.flatten().count { it == CellRole.DATA_VALUE }
        val blankCount = roleMatrix.flatten().count { it == CellRole.BLANK }

        println("=== 분류 통계 ===")
        println("HEADER: $headerCount 개")
        println("DATA_VALUE: $dataCount 개")
        println("BLANK: $blankCount 개")

        return ApiResponse(true, "성공", "분류 완료: H:$headerCount, D:$dataCount, B:$blankCount")
    }

    fun classifyFullSheet(sheet: Sheet, fullRange: Array<Int>): Array<Array<CellRole>> {
        // 결과를 저장할 2차원 배열 생성
        val roleMatrix = Array(fullRange[0] + 1) { Array(fullRange[1] + 1) { CellRole.BLANK } }

        for (i in 0..fullRange[0]) {
            for (j in 0..fullRange[1]) {
                val center = listOf(i, j)
                val nineCell = excelAnalyzeService.nineCellExtractSafe(sheet, center)
                val cellRole = classifyCellRole(nineCell)

                // 결과를 2차원 배열에 저장
                roleMatrix[i][j] = cellRole
            }
        }

        return roleMatrix
    }

    fun classifyCellRole(nineCell: Array<Array<Cell>>) : CellRole{
        val cellRole = when{
            cellClassificationService.isHeader(nineCell) -> CellRole.HEADER
            cellClassificationService.isValue(nineCell) -> CellRole.DATA_VALUE
            else -> CellRole.BLANK
        }
        return cellRole
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