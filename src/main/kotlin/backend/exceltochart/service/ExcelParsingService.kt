package backend.exceltochart.service

import backend.exceltochart.config.ApiResponse
import org.apache.commons.lang3.StringUtils.center
import org.apache.poi.ss.formula.SheetRange
import org.apache.poi.ss.usermodel.Cell
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
        val fullRange = detectSheetRange(sheet)

        // 전체 시트 분류 실행
        val roleMatrix = classifyFullSheet(sheet, fullRange)

        // 결과 출력
        printRoleMatrix(roleMatrix)

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

    // 결과 출력을 위한 헬퍼 함수
    fun printRoleMatrix(roleMatrix: Array<Array<CellRole>>) {
        println("=== Cell Role Classification Result ===")

        roleMatrix.forEach { row ->
            row.forEach { role ->
                print("[${role.name.take(6).padEnd(6)}] ")
            }
            println()
        }
    }

    enum class CellRole{
        HEADER,
        DATA_VALUE,
        CATEGORY,
        SUB_HEADER,
        BLANK,
        ROW_HEADER,
    }

    fun classifyCellRole(nineCell: Array<Array<Cell>>) : CellRole{
        val cellRole = when{
            isHeader(nineCell) -> CellRole.HEADER
            isValue(nineCell) -> CellRole.DATA_VALUE
            else -> CellRole.BLANK
        }
        return cellRole
    }

    fun isValue(nineCell: Array<Array<Cell>>): Boolean {
        val targetCell = nineCell[1][1]
        val belowCell = nineCell[2][1]  // 아래 셀
        val aboveCell = nineCell[0][1]  // 위 셀

        var score = 0

        // 1. 스타일이 bold가 아닌지 확인
        if (!isBold(targetCell)) {
            print("이거 들어가?")
            score += 1
        }

        // 2. 아래 셀과 텍스트 크기가 같은지 확인
        if (isTextSizeSame(targetCell, belowCell)) {
            score += 1
        }

        // 3. 배경색이 없거나 아래 셀과 배경색이 같은지 확인
        if (!hasBackgroundColor(targetCell) && isSameBackgroundColor(targetCell, belowCell)) {
            score += 1
        }

        // 4. 위에 셀이 없거나 위에 셀이 공백이거나 스타일이 없는지 확인
        if (!isBlankOrNoStyle(aboveCell)) {
            score += 1
        }
        print("$score 이게 점수")
        println("${getCellValue(targetCell)} 이거 뭐나오지")
        return getCellValue(targetCell) != "" && score >= 2
    }
    private fun getCellValue(cell: Cell?): String {
        return when {
            cell == null -> "null"
            cell.cellType == CellType.STRING -> cell.stringCellValue
            cell.cellType == CellType.NUMERIC -> {
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    cell.localDateTimeCellValue.toString()
                } else {
                    cell.numericCellValue.toString()
                }
            }
            cell.cellType == CellType.BOOLEAN -> cell.booleanCellValue.toString()
            cell.cellType == CellType.FORMULA -> cell.cellFormula
            else -> "empty"
        }
    }


    fun isCategory(cell:Cell): Boolean{
        return false
    }

    fun isSubHeader(cell:Cell): Boolean{
        return false
    }

    fun isBlank(cell:Cell): Boolean{
        return false
    }

    fun isRowHeader(cell:Cell): Boolean{
        return false
    }

    fun isHeader(nineCell: Array<Array<Cell>>): Boolean {
        val targetCell = nineCell[1][1]
        val belowCell = nineCell[2][1]  // 아래 셀
        val aboveCell = nineCell[0][1]  // 위 셀

        var score = 0

        // 1. 스타일이 bold인지 확인
        if (isBold(targetCell)) {
            score += 1
        }

        // 2. 아래 셀보다 텍스트 크기가 큰지 확인
        if (isTextSizeLarger(targetCell, belowCell)) {
            score += 1
        }

        // 3. 배경색이 존재하고 아래 셀과 배경색이 다른지 확인
        if (hasBackgroundColor(targetCell) && isDifferentBackgroundColor(targetCell, belowCell)) {
            score += 1
        }

        // 4. 위에 셀이 없거나 위에 셀이 공백이거나 스타일이 없는지 확인
        if (isBlankOrNoStyle(aboveCell)) {
            score += 1
        }

        return score >= 3
    }

    /**
     * 굵은 글씨가 포함되었는지 검사
     */
    private fun isBold(cell: Cell): Boolean {
        return try {
            val cellStyle = cell.cellStyle
            val font = cell.sheet.workbook.getFontAt(cellStyle.fontIndexAsInt)
            font.bold
        } catch (e: Exception) {
            false
        }
    }

    /**
     * 아래셀 기준으로 텍스트 크기 검사
     */
    private fun isTextSizeLarger(targetCell: Cell, compareCell: Cell): Boolean {
        return try {
            val targetFont = targetCell.sheet.workbook.getFontAt(targetCell.cellStyle.fontIndexAsInt)
            val compareFont = compareCell.sheet.workbook.getFontAt(compareCell.cellStyle.fontIndexAsInt)
            targetFont.fontHeightInPoints > compareFont.fontHeightInPoints
        } catch (e: Exception) {
            false
        }
    }

    private fun isTextSizeSame(targetCell: Cell, compareCell: Cell): Boolean {
        return try {
            val targetFont = targetCell.sheet.workbook.getFontAt(targetCell.cellStyle.fontIndexAsInt)
            val compareFont = compareCell.sheet.workbook.getFontAt(compareCell.cellStyle.fontIndexAsInt)
            targetFont.fontHeightInPoints == compareFont.fontHeightInPoints
        } catch (e: Exception) {
            false
        }
    }


    /**
     * 셀 배경색 존재 여부 검사
     */
    private fun hasBackgroundColor(cell: Cell): Boolean {
        return try {
            val cellStyle = cell.cellStyle
            val fillPattern = cellStyle.fillPattern
            val fillForegroundColor = cellStyle.fillForegroundColor

            fillPattern != org.apache.poi.ss.usermodel.FillPatternType.NO_FILL &&
                    fillForegroundColor != org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined.AUTOMATIC.index.toShort()
        } catch (e: Exception) {
            false
        }
    }


    /**
     * 셀 배경색 비교 검사
     */
    private fun isDifferentBackgroundColor(cell1: Cell, cell2: Cell): Boolean {
        return try {
            val color1 = cell1.cellStyle.fillForegroundColor
            val color2 = cell2.cellStyle.fillForegroundColor
            color1 != color2
        } catch (e: Exception) {
            false
        }
    }

    /**
     * 셀 배경색 비교 검사
     */
    private fun isSameBackgroundColor(cell1: Cell, cell2: Cell): Boolean {
        return try {
            val color1 = cell1.cellStyle.fillForegroundColor
            val color2 = cell2.cellStyle.fillForegroundColor
            color1 == color2
        } catch (e: Exception) {
            false
        }
    }


    /**
     * 공백셀 검사
     */
    private fun isBlankOrNoStyle(cell: Cell): Boolean {
        // 셀이 공백인지 확인
        val isBlank = when (cell.cellType) {
            CellType.BLANK -> true
            CellType.STRING -> cell.stringCellValue.trim().isEmpty()
            else -> false
        }

        if (isBlank) return true

        // 스타일이 없는지 확인 (기본 스타일인지)
        return try {
            val cellStyle = cell.cellStyle
            val font = cell.sheet.workbook.getFontAt(cellStyle.fontIndexAsInt)

            // 기본 폰트 설정인지 확인
            !font.bold &&
                    font.fontHeightInPoints <= 11 &&
                    cellStyle.fillPattern == org.apache.poi.ss.usermodel.FillPatternType.NO_FILL
        } catch (e: Exception) {
            true
        }
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