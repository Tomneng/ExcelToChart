package backend.exceltochart.controller

import backend.exceltochart.config.ApiResponse
import backend.exceltochart.service.ExcelParsingService
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RequestParam
import org.springframework.web.bind.annotation.RestController
import org.springframework.web.multipart.MultipartFile

@RestController
@RequestMapping("/api")
class FileController {

    val excelParsingService = ExcelParsingService()

    @PostMapping("upload")
    fun uploadFile(@RequestParam("file") file: MultipartFile): ApiResponse {
        return excelParsingService.validateExcel(file)
    }
}