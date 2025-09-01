package backend.exceltochart.service

import backend.exceltochart.model.CellRole
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.springframework.stereotype.Service

@Service
class CellInfoUtilService {

    fun getCellValue(cell: Cell?): String {
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

    fun printRoleMatrix(roleMatrix: Array<Array<CellRole>>) {
        println("=== Cell Role Classification Result ===")

        roleMatrix.forEach { row ->
            row.forEach { role ->
                print("[${role.name.take(6).padEnd(6)}] ")
            }
            println()
        }
    }


    /**
     * 굵은 글씨가 포함되었는지 검사
     */
    fun isBold(cell: Cell): Boolean {
        return try {
            val cellStyle = cell.cellStyle
            val font = cell.sheet.workbook.getFontAt(cellStyle.fontIndex)
            font.bold
        } catch (e: Exception) {
            false
        }
    }

    /**
     * 아래셀 기준으로 텍스트 크기 검사
     */
    fun isTextSizeLarger(targetCell: Cell, compareCell: Cell): Boolean {
        return try {
            val targetFont = targetCell.sheet.workbook.getFontAt(targetCell.cellStyle.fontIndex)
            val compareFont = compareCell.sheet.workbook.getFontAt(compareCell.cellStyle.fontIndex)
            targetFont.fontHeightInPoints > compareFont.fontHeightInPoints
        } catch (e: Exception) {
            false
        }
    }

    fun isTextSizeSame(targetCell: Cell, compareCell: Cell): Boolean {
        return try {
            val targetFont = targetCell.sheet.workbook.getFontAt(targetCell.cellStyle.fontIndex)
            val compareFont = compareCell.sheet.workbook.getFontAt(compareCell.cellStyle.fontIndex)
            targetFont.fontHeightInPoints == compareFont.fontHeightInPoints
        } catch (e: Exception) {
            false
        }
    }


    /**
     * 셀 배경색 존재 여부 검사
     */
    fun hasBackgroundColor(cell: Cell): Boolean {
        return try {
            val cellStyle = cell.cellStyle
            val fillPattern = cellStyle.fillPattern
            val fillForegroundColor = cellStyle.fillForegroundColor

            fillPattern != org.apache.poi.ss.usermodel.FillPatternType.NO_FILL &&
                    fillForegroundColor != org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined.AUTOMATIC.index
        } catch (e: Exception) {
            false
        }
    }


    /**
     * 셀 배경색 비교 검사
     */
    fun isDifferentBackgroundColor(cell1: Cell, cell2: Cell): Boolean {
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
    fun isSameBackgroundColor(cell1: Cell, cell2: Cell): Boolean {
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
    fun isBlankOrNoStyle(cell: Cell): Boolean {
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
            val font = cell.sheet.workbook.getFontAt(cellStyle.fontIndex)

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
}