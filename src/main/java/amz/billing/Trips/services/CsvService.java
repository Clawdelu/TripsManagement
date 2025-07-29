package amz.billing.Trips.services;

import amz.billing.Trips.entities.Stop;
import amz.billing.Trips.entities.Trip;
import amz.billing.Trips.enums.Status;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import com.opencsv.exceptions.CsvValidationException;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

@RequiredArgsConstructor
@Service
public class CsvService implements ICsvService {

    private final IExcelService excelService;

    @Override
    public List<Trip> processCsvFile(MultipartFile csvReadFromFile) {
        List<Trip> trips = new ArrayList<>();
        try (CSVReader reader = new CSVReader(new InputStreamReader(csvReadFromFile.getInputStream()))) {
            String[] line;
            reader.skip(1);

            while ((line = reader.readNext()) != null) {
                Status status;
                try {
                    status = Status.valueOf(line[4].toUpperCase());
                } catch (IllegalArgumentException | NullPointerException e) {
                    status = Status.COMPLETED;
                }
                Trip trip = Trip.builder()
                        .vrid(line[2])
                        .price(Double.parseDouble(line[15].split(" ")[0]))
                        .status(status)
                        .driverList(Stream.of(line[17].split(",")).collect(Collectors.toList()))
                        .trailerID(!line[14].isEmpty() ?
                                line[14].indexOf("-") == line[14].lastIndexOf('-') ? line[14].split("-")[1] : line[14]
                                : line[14])
                        .vehicleID(!line[18].isEmpty() ?
                                line[18].indexOf('-') == line[18].lastIndexOf('-') ? line[18].split("-")[1] : line[18] :
                                line[18])
                        .transitOperatorType(line[51].toUpperCase())
                        .build();

                if(trip.getStatus().equals(Status.CANCELLED) ){
                    double actualCancelPrice = 220;
                    if(trip.getTransitOperatorType().equals("TEAM_DRIVER"))
                        actualCancelPrice = 320;
                    if(!trip.getPrice().equals((double) 0)) {
                        System.out.println("Schimbat pret la: " + trip.getVrid() + "de la " + trip.getPrice() + " in " + actualCancelPrice);
                        trip.setPrice(actualCancelPrice);
                    }
                }

                String[] finalLine = line;
                List<Stop> stopList = IntStream.iterate(19, i -> i + 3)
                        .limit(10)
                        .filter(i -> !finalLine[i].isEmpty())
                        .mapToObj(i -> createStop(finalLine, i))
                        .collect(Collectors.toList());

                trip.setStopList(stopList);


//                List<Stop> stopList= new ArrayList<>();
//                for (int i = 19; i <= 48; i+=3){
//                    if(!line[i].isEmpty()){
//                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM-dd-yy HH:mm");
//                        stopList.add(Stop.builder()
//                                        .stopName(line[i])
//                                        .stopYardArrival(LocalDateTime.parse(line[i+1], formatter))
//                                        .stopYardDeparture(LocalDateTime.parse(line[i+2], formatter))
//                                .build());
//                    }
//                }
//
//                trip.setStopList(stopList);
                trips.add(trip);
            }
        } catch (IOException | CsvValidationException e) {
            throw new RuntimeException("Nu s-a putut deschide fisierul .csv");
        }
        return trips;
    }

    @Override
    public void writeTripsToCsv(String filePath, List<Trip> trips) {
        try (CSVWriter writer = new CSVWriter(new FileWriter(filePath))) {
            writer.writeNext(new String[]{"VRID", "Price", "Status", "TrailerID", "VehicleID", "Driver", "Stop 1", "Stop 2", "Stop 3"});

            for (Trip trip : trips) {
                writer.writeNext(new String[]{
                        trip.getVrid(),
                        String.valueOf(trip.getPrice()),
                        trip.getStatus().name(),
                        trip.getTrailerID(),
                        trip.getVehicleID(),
                        trip.getDriverList().toString(),
                        trip.getStopList().get(0).getStopName(),
                        trip.getStopList().get(1).getStopName(),
                        trip.getStopList().size() > 2 ? trip.getStopList().get(2).getStopName() : "",
                });
            }
        } catch (IOException e) {
            throw new RuntimeException("Nu am putut citi fisierul CSV.");
        }
    }

    @Override
    public byte[] exportNewExcelFile(MultipartFile csvFile) {
        List<Trip> tripList = processCsvFile(csvFile).stream()
                .sorted(Comparator.comparing(Trip::getVehicleID)
                        .thenComparing(trip -> trip.getStopList().isEmpty() ? LocalDateTime.MAX :
                                trip.getStopList().get(0).getStopYardArrival()))
                .collect(Collectors.toList());
        return excelService.createNewExcel(tripList);
    }

    @Override
    public byte[] exportSameExcelFile(MultipartFile csvFile, MultipartFile excelFile) {


        List<Trip> tripList = processCsvFile(csvFile).stream()
                .sorted(Comparator.comparing(Trip::getVehicleID)
                        .thenComparing(trip -> trip.getStopList().isEmpty() ? LocalDateTime.MAX :
                                trip.getStopList().get(0).getStopYardArrival()))
                .collect(Collectors.toList());
        return excelService.updateTripsInExcel(excelFile, tripList);
    }

    private static Stop createStop(String[] line, int index) {
        List<DateTimeFormatter> formatters = Arrays.asList(
                DateTimeFormatter.ofPattern("MM-dd-yy H:mm"),
                DateTimeFormatter.ofPattern("MM-dd-yy HH:mm"),
                DateTimeFormatter.ofPattern("MM/dd/yyyy HH:mm"),
                DateTimeFormatter.ofPattern("MM/dd/yyyy H:mm")
        );

        LocalDateTime stopYardArrival = null;
        LocalDateTime stopYardDeparture = null;
        boolean parsedSuccessfully = false;

        for (DateTimeFormatter formatter : formatters) {
            try {
                stopYardArrival = LocalDateTime.parse(line[index + 1], formatter);
                stopYardDeparture = !line[index + 2].isEmpty() ? LocalDateTime.parse(line[index + 2], formatter) : null;
                parsedSuccessfully = true;
                break;
            } catch (Exception e) {
                //throw new RuntimeException("Nu am reusit sa parsam la csv yard arrival/departure");
            }
        }

        if (!parsedSuccessfully) {
            throw new RuntimeException("Nu am reușit să parsam la CSV yard arrival/departure pentru linia: " + Arrays.toString(line));
        }

        return Stop.builder()
                .stopName(line[index])
                .stopYardArrival(stopYardArrival)
                .stopYardDeparture(stopYardDeparture)
                .build();
    }
}
