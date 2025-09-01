package backend.exceltochart.service

import org.apache.poi.ss.usermodel.Cell
import org.springframework.stereotype.Service

@Service
class CellClassificationService(
    private val cellInfoUtilService: CellInfoUtilService
) {

    fun isHeader(nineCell: Array<Array<Cell>>): Boolean {
        val targetCell = nineCell[1][1]
        val belowCell = nineCell[2][1]  // 아래 셀
        val aboveCell = nineCell[0][1]  // 위 셀

        var score = 0

        // 1. 스타일이 bold인지 확인
        if (cellInfoUtilService.isBold(targetCell)) {
            score += 1
        }

        // 2. 아래 셀보다 텍스트 크기가 큰지 확인
        if (cellInfoUtilService.isTextSizeLarger(targetCell, belowCell)) {
            score += 1
        }

        // 3. 배경색이 존재하고 아래 셀과 배경색이 다른지 확인
        if (cellInfoUtilService.hasBackgroundColor(targetCell) && cellInfoUtilService.isDifferentBackgroundColor(targetCell, belowCell)) {
            score += 1
        }

        // 4. 위에 셀이 없거나 위에 셀이 공백이거나 스타일이 없는지 확인
        if (cellInfoUtilService.isBlankOrNoStyle(aboveCell)) {
            score += 1
        }

        return score >= 3
    }

    fun isValue(nineCell: Array<Array<Cell>>): Boolean {
        val targetCell = nineCell[1][1]
        val belowCell = nineCell[2][1]  // 아래 셀
        val aboveCell = nineCell[0][1]  // 위 셀

        var score = 0

        // 1. 스타일이 bold가 아닌지 확인
        if (!cellInfoUtilService.isBold(targetCell)) {
            print("이거 들어가?")
            score += 1
        }

        // 2. 아래 셀과 텍스트 크기가 같은지 확인
        if (cellInfoUtilService.isTextSizeSame(targetCell, belowCell)) {
            score += 1
        }

        // 3. 배경색이 없거나 아래 셀과 배경색이 같은지 확인
        if (!cellInfoUtilService.hasBackgroundColor(targetCell) && cellInfoUtilService.isSameBackgroundColor(targetCell, belowCell)) {
            score += 1
        }

        // 4. 위에 셀이 없거나 위에 셀이 공백이거나 스타일이 없는지 확인
        if (!cellInfoUtilService.isBlankOrNoStyle(aboveCell)) {
            score += 1
        }

        return cellInfoUtilService.getCellValue(targetCell) != "" && score >= 2
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


}