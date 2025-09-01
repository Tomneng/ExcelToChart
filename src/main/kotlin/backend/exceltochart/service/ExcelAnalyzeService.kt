package backend.exceltochart.service

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.springframework.stereotype.Service

@Service
class ExcelAnalyzeService {

    enum class CellRole{
        HEADER,
        VALUE,
        SUB_HEADER,
    }


    private fun createEmptyCell(sheet: Sheet): Cell {
        // 임시 행과 셀 생성 (실제 시트에 추가되지 않음)
        val tempRow = sheet.createRow(sheet.lastRowNum + 1)
        val emptyCell = tempRow.createCell(0)
        emptyCell.setCellValue("")
        return emptyCell
    }

    // 대안: 실제로 시트에 빈 셀을 생성하지 않고 기존 빈 셀 재사용
    fun nineCellExtractSafe(sheet: Sheet, center: List<Int>): Array<Array<Cell>> {
        val nineCell: Array<Array<Cell?>> = Array(3) { arrayOfNulls(3) }
        var defaultCell: Cell? = null

        val centerRow = center[0]
        val centerCol = center[1]

        for (i in 0..2) {
            for (j in 0..2) {
                val rowIndex = centerRow - 1 + i
                val colIndex = centerCol - 1 + j

                if (rowIndex >= 0 && colIndex >= 0) {
                    val row = sheet.getRow(rowIndex)
                    if (row != null && colIndex < row.lastCellNum) {
                        val cell = row.getCell(colIndex)
                        nineCell[i][j] = cell

                        // 첫 번째로 찾은 빈 셀을 기본값으로 사용
                        if (defaultCell == null && cell != null && getCellValue(cell).trim().isEmpty()) {
                            defaultCell = cell
                        }
                    }
                }
            }
        }

        // 기본 셀이 없으면 새로 생성
        if (defaultCell == null) {
            defaultCell = createEmptyCell(sheet)
        }

        nineCell.forEach { row ->
            row.forEach { cell ->
                print("[${getCellValue(cell)}] ")
            }
            println()
        }

        val finalDefaultCell = defaultCell
        return Array(3) { i ->
            Array(3) { j ->
                nineCell[i][j] ?: finalDefaultCell
            }
        }
    }
    // 3. 헬퍼 함수 사용
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

//    private fun classifyCell(cell: List<Cell>): CellRole{
//
//
//    }

}