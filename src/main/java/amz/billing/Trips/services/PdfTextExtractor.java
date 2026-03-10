package amz.billing.Trips.services;

import amz.billing.Trips.exception.DocumentExtractionException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.springframework.stereotype.Service;

@Service
@Slf4j
public class PdfTextExtractor implements DocumentTextExtractor {

  public String extractText(InputStream documentStream) {

    Path tempFilePath = null;
    try {
      tempFilePath = Files.createTempFile("upload_pdf_", ".pdf");
      Files.copy(documentStream, tempFilePath, StandardCopyOption.REPLACE_EXISTING);
      File pdfFile = tempFilePath.toFile();

      try (PDDocument document = Loader.loadPDF(pdfFile)) {
        PDFTextStripper stripper = new PDFTextStripper();
        stripper.setSortByPosition(true);
        return stripper.getText(document);
      }

    } catch (IOException e) {
      throw new DocumentExtractionException("An error occurred during PDF text extraction.", e);
    } finally {
      if (tempFilePath != null) {
        try {
          Files.deleteIfExists(tempFilePath);
        } catch (IOException ignored) {
          log.warn("Could not delete the temp file.", ignored);
        }
      }
    }
  }
}
