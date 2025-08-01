package amz.billing.Trips.services;

import amz.billing.Trips.entities.Company;
import amz.billing.Trips.entities.Stop;
import amz.billing.Trips.entities.Trip;
import amz.billing.Trips.enums.Status;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.WeekFields;
import java.util.*;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RequiredArgsConstructor
@Service
public class ExcelService implements IExcelService {
    public static final List<Company> companyList = List.of(
            new Company(UUID.randomUUID().toString(), "PITAR", List.of("SV01PIT","SV03PIT","SV04PTI","SV08PTI","SV18PIT","SV28PIT","SV31PIT","SV32PIT","SV37PIT","SV41PIT",
                    "SV62PYT","SV64PIT","SV71PIT","SV83PIT","SV86PIT"), "PITAR", 3, 4.5),
            new Company(UUID.randomUUID().toString(), "PHT", List.of("SV19PHT", "SV23SPD", "SV24SPD", "SV44PHT","SV48PHT","SV52PHT", "SV55PHT", "SV77PHT", "SV93PHT"), "PHT", 2, 3.5),
            new Company(UUID.randomUUID().toString(), "Prime", List.of("2GFP388", "2GXY352"), "PRIME NOVA TRANS", 0, 5),
            new Company(UUID.randomUUID().toString(), "Royal", List.of("SV16FHK", "SV40RMT"), "Royal", 5, 5),
            new Company(UUID.randomUUID().toString(), "AutoLuk", List.of("B28SPD"), "AUTOLUK", 9, 9),
            new Company(UUID.randomUUID().toString(), "Nilo", List.of("SV38DUC"), "NILO", 9, 9),
            new Company(UUID.randomUUID().toString(), "Eurot", List.of("SV39EUR", "SV50EUR", "SV62EUR"), "EUROTRANSFLOR", 9, 9),
            new Company(UUID.randomUUID().toString(), "Miniflor", List.of("SV05MNY", "SV65MYN", "SV11MNI", "SV48MYM"), "Miniflor", 9, 9),
            new Company(UUID.randomUUID().toString(), "Cupola", List.of("SV26CUP"), "CUP", 9, 9),
            new Company(UUID.randomUUID().toString(), "DUO", List.of("B999DKL"), "DUO Kanlogistik", 9, 9),
            new Company(UUID.randomUUID().toString(), "Bucovina", List.of("B79CBT"), "CRISPOP BUCOVINA TRANSPORT", 9, 9),
            new Company(UUID.randomUUID().toString(), "Brumar", List.of("BC34BRU", "BC95BRU","BC99BRU"), "BRUMAR CARM SRL", 9, 9),
            new Company(UUID.randomUUID().toString(), "Sabdary", List.of("SV10SDT", "SV26SDT", "SV76SDT"), "Sabdary", 9, 9),
            new Company(UUID.randomUUID().toString(), "Johandav", List.of("SV01JHH", "SV02JHH", "SV04JHH", "SV05JHH", "SV06JHH", "SV08JHH", "SV09JHH", "SV10JHH", "SV12JHH", "SV18JHH",
                    "SV19JHH", "SV20JHH", "SV22JHH", "SV24JHH", "SV26JHH", "SV28JHH", "SV30JHH", "SV31JHH", "SV32JHH", "SV33JHH", "SV34JHH", "SV35JHH", "SV36JHH", "SV37JHH", "SV38JHH","SV39JHH","SV41JHH", "SV95NMD"), "JOHANDAV", 5, 4.25),
            new Company(UUID.randomUUID().toString(), "LLS", List.of("SV96LLS"), "LLT", 3, 5),
            new Company(UUID.randomUUID().toString(), "Farlan", List.of("SV39FAR","SV57FAR","SV78FAR","SV29VNC","SV74FRL","SV47FRL"), "FARLAN", 10, 3.5),
            new Company(UUID.randomUUID().toString(), "HIGHWAY", List.of("NT77LKW"), "HIGHWAY TRUCKS SRL", 9, 9),
            new Company(UUID.randomUUID().toString(), "Fabian Truck", List.of("B250MMX", "B251MMX", "B252MMX", "B253MMX", "B254MMX", "B255MMX", "B256MMX", "SV18MMX","SV89MMX"), "FABIAN TRUCK SRL", 5, 4.25),
            new Company(UUID.randomUUID().toString(), "OTHER", List.of(""), "Other SRL", 5, 5)

    );

    @Override
    public byte[] updateTripsInExcel(MultipartFile excelFile, List<Trip> tripList) {
        try (XSSFWorkbook xssfWorkbook = new XSSFWorkbook(excelFile.getInputStream())) {
            // SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook);
            Map<String, List<Trip>> groupedByVehicleID = tripList.stream()
                    .collect(Collectors.groupingBy(Trip::getVehicleID, LinkedHashMap::new, Collectors.toList()));

            for (String key : groupedByVehicleID.keySet()) {
                Sheet sheet = xssfWorkbook.getSheet(key);
                if (sheet == null) {
                    System.out.println("Sheet-ul " + key + " nu există.");
                    sheet = xssfWorkbook.getSheet("Others");
//                    if (sheet == null) {
//                        sheet = sxssfWorkbook.createSheet("Others");
//                        createTable(sxssfWorkbook, sheet, groupedByVehicleID.get(key));
//                    }else {
//                        updateTable(sheet,groupedByVehicleID.get(key),sxssfWorkbook);
//                    }
                    updateTable(sheet, groupedByVehicleID.get(key), xssfWorkbook);
                } else {
                    System.out.println("Sheet găsit: " + key);
                    updateTable(sheet, groupedByVehicleID.get(key), xssfWorkbook);
                }
            }

//            for (var trip : tripList) {
//                Sheet sheet = sxssfWorkbook.getSheet(trip.getVehicleID());
//                if (sheet != null) {
//                    System.out.println("Sheet găsit !" + trip.getVehicleID());
//                    updateTable(sheet, trip, sxssfWorkbook);
//                } else {
//                    System.out.println("Sheet-ul " + trip.getVehicleID() + " nu există.");
//                }
//            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
//            sxssfWorkbook.write(outputStream);
//            sxssfWorkbook.dispose();
//            sxssfWorkbook.close();
            xssfWorkbook.write(outputStream);
            xssfWorkbook.close();

            return outputStream.toByteArray();

        } catch (IOException e) {
            throw new RuntimeException("Nu s-a putut deschide fisierul .xlsx");
        }
    }

    @Override
    public byte[] createNewExcel(List<Trip> tripList) {
        try (SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(100)) {

            Map<String, List<Trip>> groupedByVehicleID = tripList.stream()
                    .collect(Collectors.groupingBy(Trip::getVehicleID, LinkedHashMap::new, Collectors.toList()));

            for (String key : groupedByVehicleID.keySet()) {
                SXSSFSheet sheet;
                if (key.isEmpty()) {
                    sheet = sxssfWorkbook.createSheet("Others");
                } else {
                    sheet = sxssfWorkbook.createSheet(key);
                }
                sheet.trackAllColumnsForAutoSizing();

                createTable(sxssfWorkbook, sheet, groupedByVehicleID.get(key));
            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            sxssfWorkbook.write(outputStream);
            sxssfWorkbook.dispose();
            sxssfWorkbook.close();
            return outputStream.toByteArray();

        } catch (IOException e) {
            throw new RuntimeException("Nu s-a putut deschide fisierul .xlsx");
        }
    }

    @Override
    public byte[] processPayment(List<MultipartFile> paymentExcelFiles, MultipartFile writeToExcel, String invoice, String payment, Integer noAnexa,
                                 String SCAC) {
        List<Trip> trips = new ArrayList<>();
        byte[] loadExcel = new byte[0];
        String gutschrift = "";
        for (MultipartFile paymentExcel : paymentExcelFiles) {

            try (InputStream inputStream = paymentExcel.getInputStream();
                 Workbook xssfWorkbook = new XSSFWorkbook(inputStream)) {

                String paymentExcelName = paymentExcel.getOriginalFilename().split("\\.")[0];

                int paymentExcelNameLength = paymentExcelName.length();
                if (paymentExcelNameLength > 8)
                    gutschrift = paymentExcelName.substring(paymentExcelNameLength - 9);
                else
                    gutschrift = paymentExcelName.substring(paymentExcelNameLength - paymentExcelNameLength / 2);

                Sheet sheet = xssfWorkbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();

                if (rowIterator.hasNext()) {
                    rowIterator.next();
                }

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Trip trip = mapRowToTripObject(row);
                    if (trip != null)
                        trips.add(trip);
                }

//            Map<String, Trip> companyTrucksMapped = new HashMap<>();
//            boolean founded = false;
//            for (var trip : trips) {
//                founded = false;
//                for (var company : companyList) {
//                    for (var x : company.getTruckIdList()) {
//                        if (trip.getVehicleID().equals(x)) {
//                            companyTrucksMapped.put(company.getKey(), trip);
//                            founded = true;
//                            break;
//                        }
//                    }
//                    if(founded) break;
//                }
//                if (!founded) {
//                    companyTrucksMapped.put("OTHER", trip);
//                }
//            }
                loadExcel = findAndUpdateExcel(writeToExcel, trips, gutschrift, invoice, payment, noAnexa);
            } catch (OLE2NotOfficeXmlFileException | POIXMLException | IOException e) {
                throw new RuntimeException("Nu am putut citi fisierele xlsx.");
            }
        }
//provizoriu
        for (var trip : trips) {
            if (trip.getStatus().equals(Status.DETENTION) && trip.getVehicleID() == null) {
                trip.setVehicleID("PHT");
            }
            if(trip.getVehicleID()==null){

                System.out.println("BULITO, nu are vehicleid" + trip.getVrid());}
        }

        Map<String, List<Trip>> companyTrucksMapped = trips.stream()
                .collect(Collectors.groupingBy(
                        trip -> companyList.stream()
                                .filter(company -> company.getTruckIdList().stream()
                                        .anyMatch(truckId -> trip.getVehicleID().equals(truckId)))
                                .findFirst()
                                .map(Company::getKey)
                                .orElse("OTHER"),
                        TreeMap::new,
                        Collectors.collectingAndThen(
                                Collectors.toList(),
                                list -> list.stream()
                                        .sorted(Comparator.comparing(Trip::getVehicleID) // Prima sortare după VehicleID
                                                .thenComparing(trip -> trip.getStopList().getFirst().getStopYardArrival())) // A doua sortare după data
                                        .collect(Collectors.toList())
                        )
                ));


        byte[] paymentExportExcel = writePaymentExcel(companyTrucksMapped, noAnexa, SCAC);

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try (ZipOutputStream zipOutputStream = new ZipOutputStream(byteArrayOutputStream)) {


//            if(paymentExportExcelList.size()>1){
//                ZipEntry entry1 = new ZipEntry("Anexa ALL fct - "+gutschrift+" wk.xlsx");
//                zipOutputStream.putNextEntry(entry1);
//                zipOutputStream.write(paymentExportExcelList.get(0));
//                zipOutputStream.closeEntry();
//
//                ZipEntry entry12 = new ZipEntry("Anexa fct - "+gutschrift+" wk.xlsx");
//                zipOutputStream.putNextEntry(entry12);
//                zipOutputStream.write(paymentExportExcelList.get(1));
//                zipOutputStream.closeEntry();
//            }

                ZipEntry entry1 = new ZipEntry("Anexa fct - "+gutschrift+" wk.xlsx");
                zipOutputStream.putNextEntry(entry1);
                zipOutputStream.write(paymentExportExcel);
                zipOutputStream.closeEntry();



            ZipEntry entry2 = new ZipEntry("AMZ.25 loads - TEST.xlsx");
            zipOutputStream.putNextEntry(entry2);
            zipOutputStream.write(loadExcel);
            zipOutputStream.closeEntry();
        } catch (IOException e) {
            throw new RuntimeException("Nu pot face un zip.");
        }


        return byteArrayOutputStream.toByteArray();


    }

    private static byte[] findAndUpdateExcel(MultipartFile excelFile, List<Trip> tripList, String gutschrift, String invoice,
                                             String payment, Integer noAnexa) {
        try (InputStream inputStream = excelFile.getInputStream();
             XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {

            for (Sheet sheet : workbook) {
                System.out.println("Sheet: " + sheet.getSheetName());
                for (Row row : sheet) {
                    Cell vridCell = row.getCell(7);

                    //if (vridCell == null) System.out.println("Celula pentru VRID lipsește pe rândul");
                   // else /
                        //System.out.println("nr: "+vridCell.getRowIndex());
                        for (Trip t : tripList) {
                            if (t.getVrid().equals(vridCell.getStringCellValue())) {
                                t.setVehicleID(row.getCell(1).getStringCellValue());

                                Cell gutschriftCell = row.getCell(9);
                                if (gutschriftCell == null) gutschriftCell = row.createCell(9);
                                if (row.getCell(9).getStringCellValue().isEmpty()) {
                                    gutschriftCell.setCellValue(gutschrift);
                                }

                                if (t.getPrice().equals(row.getCell(6).getNumericCellValue())) {
                                    row.getCell(12).setCellValue("=");
                                } else if (!t.getStatus().equals(Status.DETENTION)) {
                                    row.getCell(12).setCellValue(t.getPrice());
                                }

                                if (t.getStatus().equals(Status.DETENTION) && row.getCell(3).getStringCellValue().toLowerCase().equals("caz")) {
                                    row.getCell(12).setCellValue("D - " + t.getPrice());
                                } else if (t.getStatus().equals(Status.COMPLETED)) {
                                    row.getCell(4).setCellValue(t.getPrice() / t.getTotalDistance());
                                    if (t.getTotalDistance() == null) t.setTotalDistance(0.0);
                                    row.getCell(5).setCellValue(t.getTotalDistance());
                                }

                                Cell invoiceCell = row.getCell(10);
                                if (invoiceCell == null) invoiceCell = row.createCell(10);
                                invoiceCell.setCellValue(invoice);

                                Cell paymentCell = row.getCell(11);
                                if (paymentCell == null) paymentCell = row.createCell(11);
                                paymentCell.setCellValue(payment);

                                Cell nrFctCell = row.getCell(16);
                                if (nrFctCell == null) nrFctCell = row.createCell(16);
                                nrFctCell.setCellValue(noAnexa);
                            }
                        }
                    }

            }
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            workbook.close();
            return outputStream.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException("Nu s-a putut deschide loads excel.");
        }
    }

    private static byte[] writePaymentExcel(Map<String, List<Trip>> companyTrucksMapped, Integer noAnexa, String SCAC) {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            CellStyle centerStyle = workbook.createCellStyle();
            centerStyle.setAlignment(HorizontalAlignment.CENTER);
            centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            Font titleFont = createFont(workbook, "Times New Roman", (short) 12, true);
            Font simpleFont = createFont(workbook, "Times New Roman", (short) 12, false);
            Font simple11Font = createFont(workbook, "Times New Roman", (short) 11, false);

            CellStyle titleStyle = workbook.createCellStyle();
            titleStyle.cloneStyleFrom(centerStyle);
            titleStyle.setFont(titleFont);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(titleStyle);
            headerStyle.setBorderTop(BorderStyle.MEDIUM);
            headerStyle.setBorderRight(BorderStyle.MEDIUM);
            headerStyle.setBorderBottom(BorderStyle.MEDIUM);
            headerStyle.setBorderLeft(BorderStyle.MEDIUM);
            headerStyle.setWrapText(true);

            CellStyle firmaStyle = workbook.createCellStyle();
            firmaStyle.setFont(titleFont);

            CellStyle firmaFactStyle = workbook.createCellStyle();
            firmaFactStyle.setFont(simpleFont);


            CellStyle centerSimpleFontStyle = workbook.createCellStyle();
            centerSimpleFontStyle.cloneStyleFrom(centerStyle);
            centerSimpleFontStyle.setFont(simpleFont);

            CellStyle totalStyle = workbook.createCellStyle();
            totalStyle.cloneStyleFrom(titleStyle);
            totalStyle.setAlignment(HorizontalAlignment.LEFT);


            CellStyle centerSimple11FontStyle = workbook.createCellStyle();
            centerSimple11FontStyle.setFont(simple11Font);
            centerSimple11FontStyle.setAlignment(HorizontalAlignment.CENTER);
            centerSimple11FontStyle.setBorderRight(BorderStyle.THIN);
            centerSimple11FontStyle.setBorderLeft(BorderStyle.THIN);
            centerSimple11FontStyle.setBorderBottom(BorderStyle.THIN);

            CellStyle incarcareStyle = workbook.createCellStyle();
            incarcareStyle.cloneStyleFrom(centerSimpleFontStyle);
            incarcareStyle.setBorderBottom(BorderStyle.MEDIUM);
            incarcareStyle.setBorderRight(BorderStyle.THIN);

            CellStyle descarcareStyle = workbook.createCellStyle();
            descarcareStyle.cloneStyleFrom(centerSimpleFontStyle);
            descarcareStyle.setBorderBottom(BorderStyle.MEDIUM);
            descarcareStyle.setBorderRight(BorderStyle.MEDIUM);

            CellStyle dataCellStyle = workbook.createCellStyle();
            dataCellStyle.cloneStyleFrom(centerSimple11FontStyle);
            DataFormat dataFormat = workbook.createDataFormat();
            dataCellStyle.setDataFormat(dataFormat.getFormat("dd-MM-yy"));


            CellStyle incDescContentStyle = workbook.createCellStyle();
            incDescContentStyle.cloneStyleFrom(centerSimple11FontStyle);
            incDescContentStyle.setAlignment(HorizontalAlignment.LEFT);

            CellStyle netDiscStyle = workbook.createCellStyle();
            netDiscStyle.cloneStyleFrom(centerSimple11FontStyle);
            netDiscStyle.setAlignment(HorizontalAlignment.RIGHT);
            netDiscStyle.setDataFormat(dataFormat.getFormat("###0.00"));

            CellStyle totalNumberStyle = workbook.createCellStyle();
            totalNumberStyle.cloneStyleFrom(titleStyle);
            totalNumberStyle.setAlignment(HorizontalAlignment.RIGHT);
            totalNumberStyle.setDataFormat(dataFormat.getFormat("###0.00"));
            totalNumberStyle.setBorderRight(BorderStyle.MEDIUM);
            totalNumberStyle.setBorderLeft(BorderStyle.MEDIUM);
            totalNumberStyle.setBorderBottom(BorderStyle.MEDIUM);
            totalNumberStyle.setBorderTop(BorderStyle.MEDIUM);

            CellStyle dispoStyle = workbook.createCellStyle();
            dispoStyle.setFont(simple11Font);
            dispoStyle.setBorderLeft(BorderStyle.THIN);
            dispoStyle.setBorderBottom(BorderStyle.THIN);

            CellStyle dispo25Style = workbook.createCellStyle();
            dispo25Style.setBorderBottom(BorderStyle.THIN);

            CellStyle dispo6Style = workbook.createCellStyle();
            dispo6Style.setBorderBottom(BorderStyle.THIN);
            dispo6Style.setBorderRight(BorderStyle.THIN);

            CellStyle dispo7Style = workbook.createCellStyle();
            dispo7Style.setBorderBottom(BorderStyle.THIN);
            dispo7Style.setBorderRight(BorderStyle.THIN);
            dispo7Style.setAlignment(HorizontalAlignment.RIGHT);
            dispo7Style.setDataFormat(dataFormat.getFormat("###0.00"));

            CellStyle dispo1Style = workbook.createCellStyle();
            dispo1Style.setBorderBottom(BorderStyle.THIN);
            dispo1Style.setBorderRight(BorderStyle.THIN);
            dispo1Style.setBorderLeft(BorderStyle.THIN);

            double dispecerat = 0.0;
            double totalMinus2 = 0.0;
            double totalMinus4 = 0.0;

            double lastDiscount = 10.0;
            boolean AVEKA = false;
             //List<byte[]> exportExcels = new ArrayList<>();

            if (SCAC.equals("AVEKA")) {

                AVEKA = true;
                List<String> sheetNameAvekaList = List.of("Farlan", "PHT");
                for (var sheetName : sheetNameAvekaList) {

                    Company company = companyList.stream()
                            .filter(c -> c.getKey().equals(sheetName))
                            .findFirst()
                            .orElse(
                                    companyList.stream()
                                            .filter(c -> c.getKey().equals("OTHER"))
                                            .findFirst()
                                            .orElse(null)
                            );

                    Sheet sheet = workbook.createSheet(sheetName);
                    int rowIndex = 0;

                    sheet.setColumnWidth(0, 5 * 256);
                    sheet.setColumnWidth(1, 10 * 256);
                    sheet.setColumnWidth(2, 11 * 256);
                    sheet.setColumnWidth(3, 11 * 256);
                    sheet.setColumnWidth(4, 11 * 256);
                    sheet.setColumnWidth(5, 14 * 256);
                    sheet.setColumnWidth(6, 11 * 256);
                    sheet.setColumnWidth(7, 13 * 256);


                    Row headerRow = sheet.createRow(rowIndex++);
                    Cell firmaCell = headerRow.createCell(0);
                    firmaCell.setCellValue("PHT SERVICES SRL - Lista comenzi");
                    firmaCell.setCellStyle(firmaStyle);

                    Cell firmaFactCell = headerRow.createCell(6);
                    firmaFactCell.setCellValue(company.getFullNameCompany());
                    firmaFactCell.setCellStyle(firmaFactStyle);

                    Cell titeCell = sheet.createRow(rowIndex++).createCell(0);
                    titeCell.setCellValue("Anexa " + noAnexa);
                    titeCell.setCellStyle(titleStyle);
                    sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 7));
                    sheet.addMergedRegion(new CellRangeAddress(3, 4, 0, 0));
                    sheet.addMergedRegion(new CellRangeAddress(3, 4, 1, 1));
                    sheet.addMergedRegion(new CellRangeAddress(3, 4, 2, 2));
                    sheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 4));
                    sheet.addMergedRegion(new CellRangeAddress(3, 4, 5, 5));
                    sheet.addMergedRegion(new CellRangeAddress(3, 3, 6, 7));

                    sheet.createRow(rowIndex++);

                    Row tableHeaderRow = sheet.createRow(rowIndex++);
                    Row tableHeaderRow2 = sheet.createRow(rowIndex++);
                    String[] headers = {"Nr.Crt:", "Data", "Nr. Camion", "Loc", "Incarcare-Descarcare", "VRID", "Pret", "Net", "Discount " + company.getDiscountAveka() + "%"};
                    for (int j = 0; j < headers.length; j++) {
                        Cell headerCell = tableHeaderRow.createCell(j);
                        Cell headerCell2 = tableHeaderRow2.createCell(j);
                        headerCell2.setCellStyle(headerStyle);
                        headerCell.setCellValue(headers[j]);
                        headerCell.setCellStyle(headerStyle);
                        if (j == 3) {

                            Cell incarcareCell = tableHeaderRow2.createCell(j);
                            incarcareCell.setCellValue(headers[++j].split("-")[0]);
                            incarcareCell.setCellStyle(incarcareStyle);
                            Cell descarcareCell = tableHeaderRow2.createCell(j);
                            descarcareCell.setCellValue(headers[j].split("-")[1]);
                            descarcareCell.setCellStyle(descarcareStyle);
                            Cell locCell = tableHeaderRow.createCell(j);
                            locCell.setCellStyle(headerStyle);

                        } else if (j == 6) {
                            Cell netCell = tableHeaderRow2.createCell(j);
                            netCell.setCellValue(headers[++j]);
                            netCell.setCellStyle(incarcareStyle);
                            Cell pretCell = tableHeaderRow.createCell(j);
                            pretCell.setCellStyle(headerStyle);
                            Cell discountCell = tableHeaderRow2.createCell(j);
                            discountCell.setCellValue(headers[++j]);
                            discountCell.setCellStyle(descarcareStyle);

                        }
                    }


                   // if (companyTrucksMapped.containsKey(company.getKey())) {
                        List<Trip> tripsToAdd = new ArrayList<>();
                        if (company.getKey().equals("Farlan")) {
                            tripsToAdd = companyTrucksMapped.get(company.getKey());
                        } else {
                            tripsToAdd = companyTrucksMapped.entrySet().stream()
                                    .filter(e -> !e.getKey().equals("Farlan"))
                                    .flatMap(e -> e.getValue().stream())
                                    .toList();
                        }


                        int noOfTrip = 1;
                        for (var trip : tripsToAdd) {
                            Row newTripRow = sheet.createRow(rowIndex++);
                            Cell nrCell = newTripRow.createCell(0);
                            nrCell.setCellValue(noOfTrip++);
                            nrCell.setCellStyle(centerSimple11FontStyle);

                            Cell dataCell = newTripRow.createCell(1);
                            dataCell.setCellValue(trip.getStopList().getFirst().getStopYardArrival());
                            dataCell.setCellStyle(dataCellStyle);

                            Cell camionCell = newTripRow.createCell(2);
                            camionCell.setCellValue(trip.getVehicleID());
                            camionCell.setCellStyle(centerSimple11FontStyle);

                            Cell incarcareCell = newTripRow.createCell(3);
                            incarcareCell.setCellValue(trip.getStopList().getFirst().getStopName());
                            incarcareCell.setCellStyle(incDescContentStyle);

                            Cell descarcareCell = newTripRow.createCell(4);
                            descarcareCell.setCellValue(trip.getStopList().getLast().getStopName());
                            descarcareCell.setCellStyle(incDescContentStyle);

                            Cell vridCell = newTripRow.createCell(5);
                            vridCell.setCellValue(trip.getVrid());
                            vridCell.setCellStyle(centerSimple11FontStyle);

                            Cell netCell = newTripRow.createCell(6);
                            netCell.setCellValue(trip.getPrice());
                            netCell.setCellStyle(netDiscStyle);

                            Cell discCell = newTripRow.createCell(7);
                            discCell.setCellValue(trip.getPrice() * (1 - ((double) company.getDiscountAveka() / 100)));
                            discCell.setCellStyle(netDiscStyle);


                        }

                        sheet.createRow(rowIndex++);
                        Row totalRow = sheet.createRow(rowIndex++);
                        Cell totalCell = totalRow.createCell(5);
                        totalCell.setCellValue("Total:");
                        totalCell.setCellStyle(totalStyle);

                        String totalFormula = String.format("SUM(G6:G%d)", rowIndex - 1);
                        Cell total1Cell = totalRow.createCell(6);
                        total1Cell.setCellFormula(totalFormula);
                        total1Cell.setCellStyle(totalNumberStyle);

                        String totalFormulaDiscount = String.format("SUM(H6:H%d)", rowIndex - 1);
                        Cell total2Cell = totalRow.createCell(7);
                        total2Cell.setCellFormula(totalFormulaDiscount);
                        total2Cell.setCellStyle(totalNumberStyle);
                   // }
                }

               // workbook.write(outputStream);
                //workbook.close();
                //exportExcels.add(outputStream.toByteArray());

                  //return outputStream.toByteArray();

            }

           // XSSFWorkbook workbook1 = new XSSFWorkbook();
            //outputStream = new ByteArrayOutputStream();
            for (var x : companyList) {
                Sheet sheet = workbook.createSheet(x.getKey()+"1");
                int rowIndex = 0;

                sheet.setColumnWidth(0, 5 * 256);
                sheet.setColumnWidth(1, 10 * 256);
                sheet.setColumnWidth(2, 11 * 256);
                sheet.setColumnWidth(3, 11 * 256);
                sheet.setColumnWidth(4, 11 * 256);
                sheet.setColumnWidth(5, 14 * 256);
                sheet.setColumnWidth(6, 11 * 256);
                sheet.setColumnWidth(7, 13 * 256);


                Row headerRow = sheet.createRow(rowIndex++);
                Cell firmaCell = headerRow.createCell(0);
                firmaCell.setCellValue("PHT SERVICES SRL - Lista comenzi");
                firmaCell.setCellStyle(firmaStyle);

                Cell firmaFactCell = headerRow.createCell(6);
                firmaFactCell.setCellValue(x.getFullNameCompany());
                firmaFactCell.setCellStyle(firmaFactStyle);

                Cell titeCell = sheet.createRow(rowIndex++).createCell(0);
                titeCell.setCellValue("Anexa " + noAnexa);
                titeCell.setCellStyle(titleStyle);
                sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 7));
                sheet.addMergedRegion(new CellRangeAddress(3, 4, 0, 0));
                sheet.addMergedRegion(new CellRangeAddress(3, 4, 1, 1));
                sheet.addMergedRegion(new CellRangeAddress(3, 4, 2, 2));
                sheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 4));
                sheet.addMergedRegion(new CellRangeAddress(3, 4, 5, 5));
                sheet.addMergedRegion(new CellRangeAddress(3, 3, 6, 7));

                sheet.createRow(rowIndex++);

                Row tableHeaderRow = sheet.createRow(rowIndex++);
                Row tableHeaderRow2 = sheet.createRow(rowIndex++);

                if (AVEKA) lastDiscount = x.getDiscountAveka();
                else lastDiscount = x.getDiscount();
                String[] headers = {"Nr.Crt:", "Data", "Nr. Camion", "Loc", "Incarcare-Descarcare", "VRID", "Pret", "Net", "Discount " +lastDiscount + "%"};
                for (int j = 0; j < headers.length; j++) {
                    Cell headerCell = tableHeaderRow.createCell(j);
                    Cell headerCell2 = tableHeaderRow2.createCell(j);
                    headerCell2.setCellStyle(headerStyle);
                    headerCell.setCellValue(headers[j]);
                    headerCell.setCellStyle(headerStyle);
                    if (j == 3) {

                        Cell incarcareCell = tableHeaderRow2.createCell(j);
                        incarcareCell.setCellValue(headers[++j].split("-")[0]);
                        incarcareCell.setCellStyle(incarcareStyle);
                        Cell descarcareCell = tableHeaderRow2.createCell(j);
                        descarcareCell.setCellValue(headers[j].split("-")[1]);
                        descarcareCell.setCellStyle(descarcareStyle);
                        Cell locCell = tableHeaderRow.createCell(j);
                        locCell.setCellStyle(headerStyle);

                    } else if (j == 6) {
                        Cell netCell = tableHeaderRow2.createCell(j);
                        netCell.setCellValue(headers[++j]);
                        netCell.setCellStyle(incarcareStyle);
                        Cell pretCell = tableHeaderRow.createCell(j);
                        pretCell.setCellStyle(headerStyle);
                        Cell discountCell = tableHeaderRow2.createCell(j);
                        discountCell.setCellValue(headers[++j]);
                        discountCell.setCellStyle(descarcareStyle);

                    }
                }


                if (companyTrucksMapped.containsKey(x.getKey())) {
                    List<Trip> tripsToAdd = companyTrucksMapped.get(x.getKey());
                    int noOfTrip = 1;
                    for (var trip : tripsToAdd) {
                        Row newTripRow = sheet.createRow(rowIndex++);
                        Cell nrCell = newTripRow.createCell(0);
                        nrCell.setCellValue(noOfTrip++);
                        nrCell.setCellStyle(centerSimple11FontStyle);

                        Cell dataCell = newTripRow.createCell(1);
                        dataCell.setCellValue(trip.getStopList().getFirst().getStopYardArrival());
                        dataCell.setCellStyle(dataCellStyle);

                        Cell camionCell = newTripRow.createCell(2);
                        camionCell.setCellValue(trip.getVehicleID());
                        camionCell.setCellStyle(centerSimple11FontStyle);

                        Cell incarcareCell = newTripRow.createCell(3);
                        incarcareCell.setCellValue(trip.getStopList().getFirst().getStopName());
                        incarcareCell.setCellStyle(incDescContentStyle);

                        Cell descarcareCell = newTripRow.createCell(4);
                        descarcareCell.setCellValue(trip.getStopList().getLast().getStopName());
                        descarcareCell.setCellStyle(incDescContentStyle);

                        Cell vridCell = newTripRow.createCell(5);
                        vridCell.setCellValue(trip.getVrid());
                        vridCell.setCellStyle(centerSimple11FontStyle);

                        Cell netCell = newTripRow.createCell(6);
                        netCell.setCellValue(trip.getPrice());
                        netCell.setCellStyle(netDiscStyle);


                        Cell discCell = newTripRow.createCell(7);
                        discCell.setCellValue(trip.getPrice() * (1 - ((double) lastDiscount / 100)));
                        discCell.setCellStyle(netDiscStyle);

                        if (x.getKey().equals("PITAR") || x.getKey().equals("LLS")) {
                            dispecerat += trip.getPrice() * (lastDiscount - 2) / 100;
                            totalMinus2 += trip.getPrice();
                        } else if (!x.getKey().equals("PHT") && !x.getKey().equals("Prime")) {
                            dispecerat += trip.getPrice() * (lastDiscount - 4) / 100;
                            totalMinus4 += trip.getPrice();
                        } else if (x.getKey().equals("PHT")) {
                            totalMinus2 += trip.getPrice();
                        }

                    }

                    sheet.createRow(rowIndex++);
                    Row totalRow = sheet.createRow(rowIndex++);
                    Cell totalCell = totalRow.createCell(5);
                    totalCell.setCellValue("Total:");
                    totalCell.setCellStyle(totalStyle);

                    String totalFormula = String.format("SUM(G6:G%d)", rowIndex - 1);
                    Cell total1Cell = totalRow.createCell(6);
                    total1Cell.setCellFormula(totalFormula);
                    total1Cell.setCellStyle(totalNumberStyle);

                    String totalFormulaDiscount = String.format("SUM(H6:H%d)", rowIndex - 1);
                    Cell total2Cell = totalRow.createCell(7);
                    total2Cell.setCellFormula(totalFormulaDiscount);
                    total2Cell.setCellStyle(totalNumberStyle);
                }

            }

            totalMinus4 = totalMinus4 * 0.96;
            totalMinus2 = totalMinus2 * 0.98;

            Sheet phtSheet = workbook.getSheet("PHT1");
            int noRows = phtSheet.getLastRowNum() - 1;
            if (noRows < 6) {
                noRows++;
                phtSheet.createRow(++noRows);
                noRows++;
            }

            Row lastRow = phtSheet.createRow(noRows++);
            Cell cell1dispo = lastRow.createCell(0);
            cell1dispo.setCellStyle(dispo1Style);
            Cell dispoCell = lastRow.createCell(1);
            dispoCell.setCellValue("Dispecerat PHT");
            dispoCell.setCellStyle(dispoStyle);
            for (int i = 2; i < 6; i++) {
                Cell cell = lastRow.createCell(i);
                cell.setCellStyle(dispo25Style);
            }
            Cell cell6dispo = lastRow.createCell(6);
            cell6dispo.setCellStyle(dispo7Style);
            Cell cell7dispo = lastRow.createCell(7);
            cell7dispo.setCellStyle(dispo7Style);
            cell7dispo.setCellValue(dispecerat);

            Row penultimateRow = phtSheet.getRow(noRows - 2);
            penultimateRow.createCell(8).setCellValue("2%");
            penultimateRow.createCell(9).setCellValue("4%");

            lastRow.createCell(8).setCellValue(totalMinus2);
            lastRow.createCell(9).setCellValue(totalMinus4);

            workbook.write(outputStream);
            workbook.close();
           // exportExcels.add(outputStream.toByteArray());
          //  outputStream.close();
            return outputStream.toByteArray();
          //  return exportExcels;


        } catch (IOException e) {
            throw new RuntimeException("Nu s-a putut crea un workbook payment.");
        }
    }

    private static Trip mapRowToTripObject(Row row) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy HH:mm:ss");
        String[] parts = row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().split("->");
//        if(parts.length>2) {
//            parts[1] = parts[1] + " " + parts[2];
//            parts = Arrays.copyOf(parts, parts.length - 1);
//        }
        if (parts.length > 2) {
            parts[1] = parts[2];
            parts = Arrays.copyOf(parts, parts.length - 1);
        }

        if (!row.getCell(5).getStringCellValue().isEmpty() && !row.getCell(0).getStringCellValue().isEmpty())
            return null;

        if (!row.getCell(5).getStringCellValue().isEmpty() && row.getCell(0).getStringCellValue().isEmpty()) {
            row.createCell(6).setCellValue(row.getCell(5).getStringCellValue());
            parts = new String[]{"Stop1", "Stop2"};
            return Trip.builder()
                    .vrid(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().trim())
                    .price(row.getCell(35, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getNumericCellValue())
                    .totalDistance(1.0)
                    .status(Status.valueOf(row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().split(" - ")[1].toUpperCase()))
                    .stopList(List.of(new Stop(parts[0],
                                    LocalDateTime.now(),
                                    LocalDateTime.now()),
                            new Stop(parts[1],
                                    LocalDateTime.now(),
                                    LocalDateTime.now())))
                    .build();
        }

        if (row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().trim().equals("T-116W2TN1K"))
            parts = new String[]{"Stop1", "Stop2"};


        String statusStr = row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
        Status status;
        try {
            String[] statusParts = statusStr.split(" - ");
            if (statusParts.length > 1 && !statusParts[1].isBlank()) {
                status = Status.valueOf(statusParts[1].trim().toUpperCase());
            } else {
                status = Status.COMPLETED;
            }
        } catch (Exception e) {
            status = Status.COMPLETED;
        }
        return Trip.builder()
                .vrid(row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().trim())
                .price(row.getCell(35, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getNumericCellValue())
                .totalDistance(Double.valueOf(row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().split(" ")[0]))
                //.status(Status.valueOf(row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().split(" - ")[1].toUpperCase()))
                .status(status)
                .stopList(List.of(new Stop(parts[0],
                                LocalDateTime.parse(row.getCell(7).getStringCellValue().replace(" UTC", ""), formatter),
                                LocalDateTime.now()),
                        new Stop(parts[1],
                                LocalDateTime.parse(row.getCell(7).getStringCellValue().replace(" UTC", ""), formatter),
                                LocalDateTime.now())))
                .build();
    }

    private static void updateTable(Sheet sheet, List<Trip> tripList, XSSFWorkbook xssfWorkbook) {

        Font contentFont = createFont(xssfWorkbook, "Calibri", (short) 12, true);
        Font redFont = createFont(xssfWorkbook, "Calibri", (short) 12, true, new byte[]{(byte) 192, (byte) 0, (byte) 0});

        CellStyle[] styles = {
                createCellStyle(xssfWorkbook, new byte[]{(byte) 221, (byte) 235, (byte) 247}),
                createCellStyle(xssfWorkbook, new byte[]{(byte) 252, (byte) 228, (byte) 214}),
                createCellStyle(xssfWorkbook, new byte[]{(byte) 255, (byte) 248, (byte) 235}),
                createCellStyle(xssfWorkbook, new byte[]{(byte) 237, (byte) 237, (byte) 237}),
                createCellStyle(xssfWorkbook, new byte[]{(byte) 255, (byte) 236, (byte) 197})
        };

        CellStyle contentStyle0 = createContentStyle(xssfWorkbook, contentFont, styles[0]);
        CellStyle contentStyle1 = createContentStyle(xssfWorkbook, contentFont, styles[1]);
        CellStyle contentStyle2 = createContentStyle(xssfWorkbook, contentFont, styles[2]);
        CellStyle contentStyle3 = createContentStyle(xssfWorkbook, contentFont, styles[3]);
        CellStyle contentStyle4 = createContentStyle(xssfWorkbook, contentFont, styles[4]);
        CellStyle priceStyle = createContentStyle(xssfWorkbook, redFont, contentStyle1);
        DataFormat dataFormat = xssfWorkbook.createDataFormat();
        priceStyle.setDataFormat(dataFormat.getFormat("###0.00"));
        priceStyle.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle dateStyle = xssfWorkbook.createCellStyle();
        dateStyle.cloneStyleFrom(contentStyle0);
        DataFormat format = xssfWorkbook.createDataFormat();
        dateStyle.setDataFormat(format.getFormat("dd-MMM"));

        CellStyle eur_kmStyle = xssfWorkbook.createCellStyle();
        eur_kmStyle.cloneStyleFrom(contentStyle1);
        eur_kmStyle.setDataFormat(dataFormat.getFormat("###0.00"));
        eur_kmStyle.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle deFctStyle = xssfWorkbook.createCellStyle();
        deFctStyle.cloneStyleFrom(contentStyle4);
        deFctStyle.setDataFormat(dataFormat.getFormat("###0.00"));
        deFctStyle.setAlignment(HorizontalAlignment.CENTER);

        CellStyle kmMapStyle = xssfWorkbook.createCellStyle();
        kmMapStyle.cloneStyleFrom(contentStyle1);
        kmMapStyle.setAlignment(HorizontalAlignment.RIGHT);

        int lastRowNum = sheet.getLastRowNum();
        //  Row headerRow = sheet.getRow(1);
        // int columnCount = (headerRow != null) ? headerRow.getPhysicalNumberOfCells() : 17;
        Set<String> existingValues = new HashSet<>();
        for (int i = 2; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                existingValues.add(cell.getStringCellValue().trim());
            }
        }

        for (Trip trip : tripList) {
            if (!existingValues.contains(trip.getVrid())) {
                Row newRow = sheet.createRow(++lastRowNum);
                printToCell(newRow, trip, lastRowNum + 1);
                for (int j = 0; j < 17; j++) {
                    if (newRow.getCell(j) == null) newRow.createCell(j);
                    if (j == 6)
                        newRow.getCell(j).setCellStyle(priceStyle);
                    else if (j == 0)
                        newRow.getCell(j).setCellStyle(dateStyle);
                    else if (j == 4)
                        newRow.getCell(j).setCellStyle(eur_kmStyle);
                    else if (j == 5)
                        newRow.getCell(j).setCellStyle(kmMapStyle);
                    else if (j == 15)
                        newRow.getCell(j).setCellStyle(deFctStyle);
                    else if (j < 4)
                        newRow.getCell(j).setCellStyle(contentStyle0);
                    else if (j < 12)
                        newRow.getCell(j).setCellStyle(contentStyle2);
                    else if (j < 15)
                        newRow.getCell(j).setCellStyle(contentStyle3);
                    else
                        newRow.getCell(j).setCellStyle(contentStyle4);
                }
                System.out.println("Adaugat: " + trip.getVrid());
            } else {
                System.out.println("VRID este deja: " + trip.getVrid());
            }
        }

    }

    private static void createTable(SXSSFWorkbook workbook, Sheet sheet, List<Trip> trips) {
        int rowIndex = 0;
        sheet.setColumnWidth(0, 12 * 256);
        sheet.setColumnWidth(1, 12 * 256);
        sheet.setColumnWidth(2, 12 * 256);
        sheet.setColumnWidth(3, 12 * 256);
        sheet.setColumnWidth(4, 12 * 256);
        sheet.setColumnWidth(5, 12 * 256);
        sheet.setColumnWidth(6, 12 * 256);
        sheet.setColumnWidth(7, 16 * 256);
        sheet.setColumnWidth(8, 8 * 256);
        sheet.setColumnWidth(9, 16 * 256);
        sheet.setColumnWidth(10, 9 * 256);
        sheet.setColumnWidth(11, 9 * 256);
        sheet.setColumnWidth(12, 8 * 256);
        sheet.setColumnWidth(13, 28 * 256);
        sheet.setColumnWidth(14, 8 * 256);
        sheet.setColumnWidth(15, 10 * 256);
        sheet.setColumnWidth(16, 10 * 256);

        Row firstRow = sheet.createRow(rowIndex++);

        int[][] mergedRegions = {
                {0, 0, 0, 3},  // LOADS / UNLOADS
                {0, 0, 4, 6},  // PRICE
                {0, 0, 7, 11}, // DOCUMENTS / PAYMENT
                {0, 0, 12, 14}, // OBSERVATII
                {0, 0, 15, 16}  // Last two columns
        };

        for (int[] region : mergedRegions) {
            sheet.addMergedRegion(new CellRangeAddress(region[0], region[1], region[2], region[3]));
        }

        Font titleFont = createFont(workbook, "Arial Black", (short) 12, true);
        Font contentFont = createFont(workbook, "Calibri", (short) 12, true);
        Font redFont = createFont(workbook, "Calibri", (short) 12, true, new byte[]{(byte) 192, (byte) 0, (byte) 0});

        CellStyle[] styles = {
                createCellStyle(workbook, new byte[]{(byte) 221, (byte) 235, (byte) 247}),
                createCellStyle(workbook, new byte[]{(byte) 252, (byte) 228, (byte) 214}),
                createCellStyle(workbook, new byte[]{(byte) 255, (byte) 248, (byte) 235}),
                createCellStyle(workbook, new byte[]{(byte) 237, (byte) 237, (byte) 237}),
                createCellStyle(workbook, new byte[]{(byte) 255, (byte) 236, (byte) 197})
        };

        CellStyle[] headerStyles = new CellStyle[5];
        for (int i = 0; i < styles.length; i++) {
            headerStyles[i] = cloneCellStyle(workbook, styles[i], titleFont, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        }

        createMergedCell(firstRow, 0, 4, "LOADS / UNLOADS", headerStyles[0]);
        createMergedCell(firstRow, 4, 7, "PRICE", headerStyles[1]);
        createMergedCell(firstRow, 7, 12, "DOCUMENTS / PAYMENT", headerStyles[2]);
        createMergedCell(firstRow, 12, 15, "OBSERVATII", headerStyles[3]);
        createMergedCell(firstRow, 15, 17, "", headerStyles[4]);

        String[] headers = {"Date", "No.Auto", "Load", "Unload", "EUR/KM", "Km Map", "Price", "Trip", "Week", "Gutschrift", "Invoice", "Payment", "", "", "", "de fct", "nr fct"};
        Row headerRow = sheet.createRow(rowIndex++);

        CellStyle contentStyle0 = createContentStyle(workbook, contentFont, styles[0]);
        CellStyle contentStyle1 = createContentStyle(workbook, contentFont, styles[1]);
        CellStyle contentStyle2 = createContentStyle(workbook, contentFont, styles[2]);
        CellStyle contentStyle3 = createContentStyle(workbook, contentFont, styles[3]);
        CellStyle contentStyle4 = createContentStyle(workbook, contentFont, styles[4]);
        CellStyle priceStyle = createContentStyle(workbook, redFont, contentStyle1);
        DataFormat dataFormat = workbook.createDataFormat();
        priceStyle.setDataFormat(dataFormat.getFormat("###0.00"));
        priceStyle.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.cloneStyleFrom(contentStyle0);
        DataFormat format = workbook.createDataFormat();
        dateStyle.setDataFormat(format.getFormat("dd-MMM"));

        CellStyle eur_kmStyle = workbook.createCellStyle();
        eur_kmStyle.cloneStyleFrom(contentStyle1);
        eur_kmStyle.setDataFormat(dataFormat.getFormat("###0.00"));
        eur_kmStyle.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle kmMapStyle = workbook.createCellStyle();
        kmMapStyle.cloneStyleFrom(contentStyle1);
        kmMapStyle.setAlignment(HorizontalAlignment.RIGHT);


        for (int j = 0; j < headers.length; j++) {
            Cell cell = headerRow.createCell(j);
            if (!headers[j].isEmpty())
                cell.setCellValue(headers[j]);
            if (j < 4) {
                cell.setCellStyle(contentStyle0);
            } else if (j < 7) {
                cell.setCellStyle(contentStyle1);
            } else if (j < 12) {
                cell.setCellStyle(contentStyle2);
            } else if (j < 15) {
                cell.setCellStyle(contentStyle3);
            } else {
                cell.setCellStyle(contentStyle4);
            }
        }
        Row emptyRow = sheet.createRow(rowIndex++);
        for (int j = 0; j < headers.length; j++) {
            Cell cell = emptyRow.createCell(j);
            if (j < 4) {
                cell.setCellStyle(contentStyle0);
            } else if (j < 7) {
                cell.setCellStyle(contentStyle1);
            } else if (j < 12) {
                cell.setCellStyle(contentStyle2);
            } else if (j < 15) {
                cell.setCellStyle(contentStyle3);
            } else {
                cell.setCellStyle(contentStyle4);
            }

        }

        for (Trip trip : trips) {
            Row row = sheet.createRow(rowIndex++);
            printToCell(row, trip, rowIndex);
            for (int j = 0; j < headers.length; j++) {
                if (row.getCell(j) == null) row.createCell(j);
                if (j == 6)
                    row.getCell(j).setCellStyle(priceStyle);
                else if (j == 0)
                    row.getCell(j).setCellStyle(dateStyle);
                else if (j == 4)
                    row.getCell(j).setCellStyle(eur_kmStyle);
                else if (j == 5)
                    row.getCell(j).setCellStyle(kmMapStyle);
                else if (j < 4)
                    row.getCell(j).setCellStyle(contentStyle0);
                else if (j < 12)
                    row.getCell(j).setCellStyle(contentStyle2);
                else if (j < 15)
                    row.getCell(j).setCellStyle(contentStyle3);
                else
                    row.getCell(j).setCellStyle(contentStyle4);
            }
        }
    }

    private static Font createFont(SXSSFWorkbook workbook, String name, short size, boolean bold) {
        Font font = workbook.createFont();
        font.setFontName(name);
        font.setFontHeightInPoints(size);
        font.setBold(bold);
        return font;
    }

    private static Font createFont(XSSFWorkbook workbook, String name, short size, boolean bold) {
        Font font = workbook.createFont();
        font.setFontName(name);
        font.setFontHeightInPoints(size);
        font.setBold(bold);
        return font;
    }

    private static Font createFont(SXSSFWorkbook workbook, String name, short size, boolean bold, byte[] color) {
        XSSFFont font = (XSSFFont) createFont(workbook, name, size, bold);
        font.setColor(new XSSFColor(color, null));
        return font;
    }

    private static Font createFont(XSSFWorkbook workbook, String name, short size, boolean bold, byte[] color) {
        XSSFFont font = (XSSFFont) createFont(workbook, name, size, bold);
        font.setColor(new XSSFColor(color, null));
        return font;
    }

    private static CellStyle createCellStyle(SXSSFWorkbook workbook, byte[] colorRGB) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(colorRGB, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static CellStyle createCellStyle(XSSFWorkbook workbook, byte[] colorRGB) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(colorRGB, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static CellStyle cloneCellStyle(SXSSFWorkbook workbook, CellStyle baseStyle, Font
            font, HorizontalAlignment horz, VerticalAlignment vert) {
        CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(baseStyle);
        newStyle.setFont(font);
        newStyle.setAlignment(horz);
        newStyle.setVerticalAlignment(vert);
        newStyle.setBorderTop(BorderStyle.THICK);
        newStyle.setBorderBottom(BorderStyle.THICK);
        newStyle.setBorderLeft(BorderStyle.THICK);
        newStyle.setBorderRight(BorderStyle.THICK);
        return newStyle;
    }

    private static CellStyle createContentStyle(SXSSFWorkbook workbook, Font font, CellStyle colorStyle) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(colorStyle);
        style.setFont(font);
        style.setBorderTop(BorderStyle.DOTTED);
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBorderLeft(BorderStyle.DOTTED);
        style.setBorderRight(BorderStyle.DOTTED);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static CellStyle createContentStyle(XSSFWorkbook workbook, Font font, CellStyle colorStyle) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(colorStyle);
        style.setFont(font);
        style.setBorderTop(BorderStyle.DOTTED);
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBorderLeft(BorderStyle.DOTTED);
        style.setBorderRight(BorderStyle.DOTTED);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static CellStyle createContentStyle(SXSSFWorkbook workbook, Font font) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setBorderTop(BorderStyle.DOTTED);
        style.setBorderBottom(BorderStyle.DOTTED);
        style.setBorderLeft(BorderStyle.DOTTED);
        style.setBorderRight(BorderStyle.DOTTED);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static void createMergedCell(Row row, int column, int toC, String text, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        for (int i = column + 1; i < toC; i++) {
            row.createCell(i).setCellStyle(style);
        }
    }

    private static void createI9Cell(XSSFWorkbook xssfWorkbook, Row row, String gutschrift) {
        Cell gutschriftCell = row.createCell(9);
        gutschriftCell.setCellValue(gutschrift);

        CellStyle style2 = createCellStyle(xssfWorkbook, new byte[]{(byte) 255, (byte) 248, (byte) 235});
        Font contentFont = createFont(xssfWorkbook, "Calibri", (short) 12, true);
        CellStyle contentStyle2 = createContentStyle(xssfWorkbook, contentFont, style2);

        gutschriftCell.setCellStyle(contentStyle2);
    }

    private static void printToCell(Row row, Trip trip, int rowIndex) {
        row.createCell(0).setCellValue(trip.getStopList().get(0).getStopYardArrival());
        row.createCell(1).setCellValue(trip.getVehicleID());
        var firstStop = trip.getStopList().getFirst().getStopName();
        var lastStop = trip.getStopList().getLast().getStopName();
        if (firstStop.contains("-"))
            firstStop = firstStop.split("-")[0];
        if (lastStop.contains("-"))
            lastStop = lastStop.split("-")[0];
        row.createCell(2).setCellValue(firstStop);
        row.createCell(3).setCellValue(lastStop);

        Cell cell4 = row.createCell(4);
        String formula = String.format("IFERROR(G%d/F%d, \"\")", rowIndex, rowIndex);
        cell4.setCellFormula(formula);
//        if (trip.getPrice() != null && trip.getTotalDistance() != null) {
//            cell4.setCellValue(String.format("%.2f", trip.getPrice() / trip.getTotalDistance()));
//        }
        Cell cell5 = row.createCell(5);
        if (trip.getTotalDistance() != null) {
            cell5.setCellValue(trip.getTotalDistance());
        }

        row.createCell(6).setCellValue(trip.getPrice());
        row.createCell(7).setCellValue(trip.getVrid());
        row.createCell(8).setCellValue(trip.getStopList().get(0).getStopYardArrival().get(
                WeekFields.of(Locale.getDefault()).weekOfYear()
        ));
        if (trip.getStatus().equals(Status.CANCELLED)) {
            double actualCancelPrice = 220;
            if (trip.getTransitOperatorType().equals("TEAM_DRIVER"))
                actualCancelPrice = 320;
            //row.createCell(13).setCellValue("CANCELLED - " + actualCancelPrice + " ?");
//            if (trip.getPrice().equals((double) 0))
//                row.createCell(13).setCellValue("CANCELLED - " + actualCancelPrice + " ?");
//            else
//                row.createCell(13).setCellValue("CANCELLED");
        }
        double discount = companyList.stream()
                .filter(company -> company.getTruckIdList().contains(trip.getVehicleID()))
                .map(Company::getDiscount)
                .findFirst()
                .orElse(10.0);
        String formulaDeFct = String.format("IF(G%d*%,.2f=0, \"\",G%d*(1-%,.2f/100))", rowIndex, discount, rowIndex, discount);
        row.createCell(15).setCellFormula(formulaDeFct);
    }
}


