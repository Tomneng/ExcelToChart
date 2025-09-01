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

    fun nineCellExtractSafe(sheet: Sheet, center: Array<Int>): Array<Array<Cell?>> {
        val nineCell: Array<Array<Cell?>> = Array(3) { arrayOfNulls(3) }

        val centerRow = center[0]
        val centerCol = center[1]

        for (i in 0..2) {
            for (j in 0..2) {
                val rowIndex = centerRow - 1 + i
                val colIndex = centerCol - 1 + j

                // 경계 체크
                if (rowIndex >= 0 && colIndex >= 0) {
                    val row = sheet.getRow(rowIndex)
                    if (row != null && colIndex <= row.lastCellNum) {
                        nineCell[i][j] = row.getCell(colIndex)
                    }
                }
            }
        }

        return nineCell
    }

//    private fun classifyCell(cell: List<Cell>): CellRole{
//
//
//    }

}