package amz.billing.Trips.services;

import amz.billing.Trips.entities.DKV;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@RequiredArgsConstructor
@Service
public class DkvService implements IDkvService {

    @Override
    public byte[] processDetailedDkvRaport(MultipartFile dkvRaport) {

        try (InputStream inputStream = dkvRaport.getInputStream(); XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream)) {
            Map<String, List<DKV>> dkvMapByVehicle = new HashMap<>();
            Sheet sheet = xssfWorkbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            if (rowIterator.hasNext()) rowIterator.next();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                DKV dkv = mapRowToDkvObject(row);
                if (dkv != null)
                    dkvMapByVehicle.computeIfAbsent(dkv.getVehicle(), k -> new ArrayList<>()).add(dkv);
            }

            Map<String, List<Double>> finalTotalValue = new HashMap<>();

            int countVehicles = 0;
            for (Map.Entry<String, List<DKV>> entry : dkvMapByVehicle.entrySet()) {
                String vehicle = entry.getKey();
                List<DKV> dkvList = entry.getValue();

                System.out.println("Vehicle: " + vehicle);

                double totalEur = 0.0;
                Double totalNetEur = 0.0;
                Double totalVatEur = 0.0;
                Double totalNetRon = 0.0;
                for (DKV dkv : dkvList) {
                    System.out.println("  - " + dkv.getProductName() + dkv.getTotalNetInvoiced());
                    if (dkv.getCurrency().equals("EUR")) {
                        totalNetEur += dkv.getTotalNetInvoiced();
                        totalVatEur += dkv.getTotalVatInvoiced();
                    } else {
                        totalNetRon += dkv.getTotalNetInvoiced();
                    }

                }
                totalEur += totalNetEur + totalVatEur * 0.10;


                finalTotalValue.put(vehicle, List.of(totalEur, totalNetRon));
                if (!vehicle.contains("REZ") && !vehicle.equals("empty"))
                    countVehicles++;
            }

            Double valueToSplit = finalTotalValue.get("empty").get(0) / countVehicles;
            finalTotalValue.remove("empty");

            for (String key : finalTotalValue.keySet()) {

                if (key.contains("REZ")) {
                    finalTotalValue.put(key, List.of(
                            Double.parseDouble(String.format("%.2f", finalTotalValue.get(key).get(0))),
                            Double.parseDouble(String.format("%.2f", finalTotalValue.get(key).get(1)))
                    ));
                } else {
                    finalTotalValue.put(key, List.of(
                            Double.parseDouble(String.format("%.2f", finalTotalValue.get(key).get(0) + valueToSplit)),
                            Double.parseDouble(String.format("%.2f", finalTotalValue.get(key).get(1)))
                    ));
                }

            }
            return writeToByteArray(finalTotalValue);

        } catch (OLE2NotOfficeXmlFileException | POIXMLException | IOException e) {
            throw new RuntimeException("Nu am putut citi fisierele xlsx.");
        }
    }

    private static DKV mapRowToDkvObject(Row row) {
        List<String> importantProductIdList = Arrays.asList("0949", "007J", "007N", "007S");
        if (!row.getCell(5).getStringCellValue().isEmpty() && !row.getCell(10).getStringCellValue().equals("005G")) {
            return DKV.builder()
                    .customerId(String.valueOf(row.getCell(0).getNumericCellValue()))
                    .customerName(row.getCell(1).getStringCellValue())
                    .invoiceDate(row.getCell(2).getLocalDateTimeCellValue().toLocalDate())
                    .country(row.getCell(3).getStringCellValue())
                    .cardNumber(row.getCell(4).getStringCellValue())
                    .vehicle(row.getCell(5).getStringCellValue())
                    .productId(row.getCell(10).getStringCellValue())
                    .productName(row.getCell(11).getStringCellValue())
                    .totalNetInvoiced(row.getCell(24).getNumericCellValue())
                    .totalVatInvoiced(row.getCell(25).getNumericCellValue())
                    .currency(row.getCell(23).getStringCellValue())
                    .build();
        } else if (importantProductIdList.contains(row.getCell(10).getStringCellValue())) {
            return DKV.builder()
                    .customerId(String.valueOf(row.getCell(0).getNumericCellValue()))
                    .customerName(row.getCell(1).getStringCellValue())
                    .invoiceDate(row.getCell(2).getLocalDateTimeCellValue().toLocalDate())
                    .country(row.getCell(3).getStringCellValue())
                    .cardNumber(row.getCell(4).getStringCellValue())
                    .vehicle("empty")
                    .productId(row.getCell(10).getStringCellValue())
                    .productName(row.getCell(11).getStringCellValue())
                    .totalNetInvoiced(row.getCell(24).getNumericCellValue())
                    .totalVatInvoiced(row.getCell(25).getNumericCellValue())
                    .currency(row.getCell(23).getStringCellValue())
                    .build();
        } else return null;
    }

    public static byte[] writeToByteArray(Map<String, List<Double>> map) {
        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            int count = 0;
            List<Map.Entry<String, List<Double>>> list = new ArrayList<>(map.entrySet());
            list.sort((entry1, entry2) -> {
                int keyComparison = entry1.getKey().compareTo(entry2.getKey());
                if (keyComparison != 0) {
                    return keyComparison;
                } else {
                    return entry1.getValue().get(0).compareTo(entry2.getValue().get(0));
                }
            });

            for (Map.Entry<String, List<Double>> entry : list) {
                byteArrayOutputStream.write((++count + ". " + entry.getKey() + ":\n" + "EUR: " + entry.getValue().get(0) + "\tRON: " + entry.getValue().get(1) + "\n\n").getBytes());
            }

           /* for (Map.Entry<String, Double> entry : map.entrySet()) {
                byteArrayOutputStream.write((++count +". " +entry.getKey() + ": " + entry.getValue() + "\n").getBytes());
            }*/

            return byteArrayOutputStream.toByteArray();
        } catch (IOException e) {
            System.err.println("Eroare la scrierea în ByteArray: " + e.getMessage());
            return null;
        }
    }
}