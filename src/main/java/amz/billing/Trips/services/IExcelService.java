package amz.billing.Trips.services;

import amz.billing.Trips.entities.Trip;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

public interface IExcelService {

    byte[] updateTripsInExcel(MultipartFile excelFile, List<Trip> tripList);
    byte[] createNewExcel(List<Trip> tripList);
    byte[] processPayment(List<MultipartFile> paymentExcelFiles, MultipartFile writeToExcel, String invoice, String payment, Integer noAnexa, String SCAC);
}
