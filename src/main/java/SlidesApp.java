import com.google.api.services.drive.model.File;
import com.google.api.services.sheets.v4.model.AddChartResponse;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetResponse;
import com.google.api.services.sheets.v4.model.EmbeddedChart;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.slides.v1.model.BatchUpdatePresentationResponse;
import com.google.api.services.slides.v1.model.Presentation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import service.DriveService;
import service.SheetsService;
import service.SlidesService;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

public class SlidesApp {

    private static final Logger logger = LoggerFactory.getLogger(SlidesApp.class);

    public void run() {
        try {
            generateFiles();

        } catch (IOException e) {
            logger.error("ERROR", e);
        }
    }

    private void generateFiles() throws IOException {
        String filesName = "Creating report test";
        GoogleService googleService = new GoogleService();

        // Build a new authorized API client service.
        SlidesService slidesService = new SlidesService(googleService.getSlidesService());
        SheetsService sheetsService = new SheetsService(googleService.getSheetsService());
        DriveService driveService = new DriveService(googleService.getDriveService());

        String originTemplateId = "1fXxh-0jDfFgeAibHmOa_ohEwqs4vEpLTwZVylP5-3wU";
        File presentationCopyFile = driveService.copy(originTemplateId, filesName);

        Presentation presentation = slidesService.getPresentation(presentationCopyFile.getId());

        logger.info("Created presentation from template with ID: {}", presentation.getPresentationId());
        logger.info("https://docs.google.com/presentation/d/{}", presentation.getPresentationId());

        // Create Spreadsheet
        Spreadsheet spreadsheet = sheetsService.createSpreadsheet(filesName);
        logger.info("Created spreadSheet with ID: {}", spreadsheet.getSpreadsheetId());
        logger.info("https://docs.google.com/spreadsheets/d/{}", spreadsheet.getSpreadsheetId());

        List<com.google.api.services.sheets.v4.model.Request> sheetRequests = new ArrayList<>();
        sheetRequests.add(sheetsService.generateEmbeddedChart(spreadsheet));
        sheetRequests.addAll(sheetsService.conditionalFormatting(spreadsheet));
        BatchUpdateSpreadsheetResponse sheetsRes = sheetsService.batchUpdate(spreadsheet, sheetRequests);
        logger.info("Sheets Response: {}", sheetsRes.getReplies());

        // Slides requests
        List<com.google.api.services.slides.v1.model.Request> slidesRequests = new ArrayList<>();

        AddChartResponse addChartResponse = (AddChartResponse) (sheetsRes.getReplies().get(0).get("addChart"));
        Integer chartId = ((EmbeddedChart) addChartResponse.get("chart")).getChartId();
        logger.info("ChartID: {}", chartId);

        String tableId = UUID.randomUUID().toString();

        slidesRequests.add(slidesService.addChartToSlides(spreadsheet.getSpreadsheetId(), chartId, presentation.getSlides().get(1).getObjectId()));
        slidesRequests.add(slidesService.createTableIntoSlide(tableId, presentation.getSlides().get(2).getObjectId()));
        slidesRequests.addAll(slidesService.addDataToTable(tableId));
        slidesRequests.add(slidesService.replaceText("{{mediaGroupName}}", filesName));
        slidesRequests.add(slidesService.replaceText("{{monthlyEcmpDescription}}", "Incremento del +27% mes contra mes en Mayo y +16% en Junio"));

        BatchUpdatePresentationResponse slidesRes = slidesService.batchUpdate(presentation, slidesRequests);

        logger.info("Slides Response: {}", slidesRes.getReplies());

        String pdfName = "/Users/albertmarieges/exported.pdf";
        driveService.exportPDF(presentation.getPresentationId(), pdfName);

        logger.info("PDF generated and saved as {}", pdfName);
    }

}
