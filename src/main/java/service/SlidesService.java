package service;

import com.google.api.services.slides.v1.Slides;
import com.google.api.services.slides.v1.model.*;

import java.io.IOException;
import java.util.List;

import static java.util.Arrays.asList;

public class SlidesService {

    private final Slides service;

    public SlidesService(Slides service) {
        this.service = service;
    }

    public Presentation getPresentation(String presentationID) throws IOException {
        return service.presentations().get(presentationID).execute();
    }

    public BatchUpdatePresentationResponse batchUpdate(Presentation presentation, List<Request> slidesRequests) throws IOException {
        BatchUpdatePresentationRequest slidesReq = new BatchUpdatePresentationRequest().setRequests(slidesRequests);
        return service.presentations().batchUpdate(presentation.getPresentationId(), slidesReq).execute();
    }

    public Request addChartToSlides(String spreadsheetId, Integer chartId, String slidesPageObjectId) {
        Dimension emu4M = new Dimension().setMagnitude(8000000.0).setUnit("EMU");
        return new Request()
                .setCreateSheetsChart(new CreateSheetsChartRequest()
                        .setSpreadsheetId(spreadsheetId)
                        .setChartId(chartId)
                        .setLinkingMode("LINKED")
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slidesPageObjectId)
                                .setSize(new Size()
                                        .setHeight(emu4M)
                                        .setWidth(emu4M))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(100000.0)
                                        .setTranslateY(100000.0)
                                        .setUnit("EMU"))));
    }

    public Request createTableIntoSlide(String tableId, String slidesPageObjectId) {
        return new Request()
                .setCreateTable(new CreateTableRequest()
                        .setObjectId(tableId)
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slidesPageObjectId))
                        .setColumns(6)
                        .setRows(9));
    }

    public List<Request> addDataToTable(String tableId) {
        return asList(
                createRequestEditTableText(tableId, 0, 1, "9 Feb'17 - 14 Feb'17"),
                createRequestEditTableText(tableId, 0, 2, "16 Feb'17 27 Feb'17"),
                createRequestEditTableText(tableId, 0, 3, "1 Mar'17 - 12 Mar'17"),

                createRequestEditTableText(tableId, 1, 1, "B1 - Marfeel Touch"),
                createRequestEditTableText(tableId, 1, 2, "B2 - Marfeel K"),
                createRequestEditTableText(tableId, 1, 3, "B3 - Marfeel Touch"),
                createRequestEditTableText(tableId, 1, 4, "Var. B2-B1 %"),
                createRequestEditTableText(tableId, 1, 5, "Var. B3-B1 %"),

                createRequestEditTableText(tableId, 2, 0, "ARPU"),
                createRequestEditTableText(tableId, 2, 1, "$0.86"),
                createRequestEditTableText(tableId, 2, 2, "$0.66"),
                createRequestEditTableText(tableId, 2, 3, "$0.95"),
                createRequestEditTableText(tableId, 2, 4, "-23%"),
                createRequestEditTableText(tableId, 2, 5, "11%"),

                createRequestEditTableText(tableId, 3, 0, "Impression/V"),
                createRequestEditTableText(tableId, 3, 1, "7.04"),
                createRequestEditTableText(tableId, 3, 2, "4.50"),
                createRequestEditTableText(tableId, 3, 3, "6.47"),
                createRequestEditTableText(tableId, 3, 4, "-36%"),
                createRequestEditTableText(tableId, 3, 5, "-8%"),

                new Request().setUpdateTableCellProperties(new UpdateTableCellPropertiesRequest()
                        .setFields("tableCellBackgroundFill.solidFill.color")
                        .setObjectId(tableId)
                        .setTableRange(new TableRange()
                                .setLocation(new TableCellLocation()
                                        .setColumnIndex(1)
                                        .setRowIndex(2))
                                .setRowSpan(1)
                                .setColumnSpan(3))
                        .setTableCellProperties(new TableCellProperties()
                                .setTableCellBackgroundFill(new TableCellBackgroundFill()
                                        .setSolidFill(new SolidFill()
                                                .setColor(new OpaqueColor()
                                                        .setRgbColor(new RgbColor()
                                                                .setRed(0.133333f)
                                                                .setGreen(0.545098f)
                                                                .setBlue(0.133333f)))))))
        );
    }

    private Request createRequestEditTableText(String tableId, Integer row, Integer col, String text) {
        return new Request()
                .setInsertText(new InsertTextRequest()
                        .setObjectId(tableId)
                        .setCellLocation(new TableCellLocation()
                                .setRowIndex(row)
                                .setColumnIndex(col))
                        .setText(text));
    }

    public Request replaceText(String before, String after) throws IOException {
        return new Request()
                .setReplaceAllText(new ReplaceAllTextRequest()
                        .setContainsText(new SubstringMatchCriteria()
                                .setText(before))
                        .setReplaceText(after));
    }
}
