package amz.billing.Trips.controller;

import amz.billing.Trips.services.IExcelService;
import jakarta.servlet.http.HttpSession;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@RestController
@RequiredArgsConstructor
@RequestMapping("api/excel")
public class ExcelController {


    private final IExcelService excelService;
//    @RequestMapping(  path = "/process", method = RequestMethod.POST,consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
//    public ResponseEntity<byte[]> processExcel(@RequestParam("file") MultipartFile file) throws IOException {
//
//        return ResponseEntity.ok()
//                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=processed.xlsx")
//                .contentType(MediaType.APPLICATION_OCTET_STREAM)
//                .body(excelBytes);
//    }

    @RequestMapping(path = "/get-payment-excel", method = RequestMethod.POST, consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> paymentProcess(HttpSession session, @RequestParam("paymentExcel") List<MultipartFile> paymentExcelFiles, @RequestParam("loadsExcel") MultipartFile loadsExcel,
                                            @RequestParam String invoice, @RequestParam String payment, @RequestParam Integer noAnexa, @RequestParam(required = false) String SCAC)
            throws IOException {
        if (session.getAttribute("authenticated") == null) {
            return ResponseEntity.status(401).body("Unauthorized");
        }

        var returnedExcelFile = excelService.processPayment(paymentExcelFiles, loadsExcel, invoice, payment, noAnexa, SCAC);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=payment_files.zip")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(returnedExcelFile);

    }
}
