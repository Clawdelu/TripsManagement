package amz.billing.Trips.controller;

import amz.billing.Trips.services.ICsvService;


import amz.billing.Trips.services.IDkvService;
import jakarta.servlet.http.HttpSession;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

//@Api(tags = "CSV Upload")
@RestController
@RequiredArgsConstructor
@RequestMapping("api/csv")
public class CsvController {

    private final ICsvService csvService;
    private final IDkvService dkvService;


    // Endpoint pentru a încărca un fișier CSV dintr-un request
   // @ApiOperation(value = "Upload fișier CSV", notes = "Acesta încarcă un fișier CSV pentru procesare")
//    @Operation(
//            summary = "Upload fișier CSV",
//            description = "Permite încărcarea unui fișier CSV pentru procesare"
////           requestBody = @io.swagger.v3.oas.annotations.parameters.RequestBody(
////                    description = "Fișierul CSV de încărcat",
////                    content = @Content(mediaType = MediaType.MULTIPART_FORM_DATA_VALUE)
////            )
//    )

    @RequestMapping(  path = "/upload-h-csv-xlsx", method = RequestMethod.POST,consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> uploadCsvExcel(/* @Parameter(description = "Fișierul CSV", required = true)*/ HttpSession session,
                            @RequestParam("csvFile") MultipartFile csvFile, @RequestParam("excelFile") MultipartFile excelFile) throws IOException {
        if (session.getAttribute("authenticated") == null) {
            return ResponseEntity.status(401).body("Unauthorized");
        }
        var returnedExcelFile = csvService.exportSameExcelFile(csvFile,excelFile);
     return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=processed.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(returnedExcelFile);
    }


    @RequestMapping(  path = "/upload-h-csv", method = RequestMethod.POST,consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> uploadCsv(/* @Parameter(description = "Fișierul CSV", required = true)*/HttpSession session,
            @RequestParam("csvFile") MultipartFile csvFile) throws IOException {
        if (session.getAttribute("authenticated") == null) {
            return ResponseEntity.status(401).body("Unauthorized");
        }
        var returnedExcelFile = csvService.exportNewExcelFile(csvFile);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=new_loads2025.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(returnedExcelFile);
    }

    @RequestMapping(  path = "/dkv-reporting", method = RequestMethod.POST,consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> dkvReporting(HttpSession session,
            @RequestParam("excelDKVFile") MultipartFile excelDkvFile) throws IOException {
        if (session.getAttribute("authenticated") == null) {
            return ResponseEntity.status(401).body("Unauthorized");
        }
        var resultDkvReport = dkvService.processDetailedDkvRaport(excelDkvFile);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=result-DKV.txt")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resultDkvReport);
    }
}
