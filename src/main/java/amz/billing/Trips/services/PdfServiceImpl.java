package amz.billing.Trips.services;

import org.springframework.web.multipart.MultipartFile;

public class PdfServiceImpl implements IPdfService{
    @Override
    public byte[] processPdf(MultipartFile pdfFile) {


        return new byte[0];
    }
}
