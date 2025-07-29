package amz.billing.Trips.services;

import amz.billing.Trips.entities.Trip;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

public interface ICsvService {

    List<Trip>processCsvFile(MultipartFile csvReadFromFile);
    void writeTripsToCsv(String filePath, List<Trip> trips);
    byte[] exportNewExcelFile(MultipartFile csvFile);

    byte[] exportSameExcelFile(MultipartFile csvFile, MultipartFile excelFile);
}
