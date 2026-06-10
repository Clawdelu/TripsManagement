package amz.billing.Trips.services;

import org.springframework.web.multipart.MultipartFile;

public interface IPdfService {

  byte[] processPdf(MultipartFile pdfFile);
}
