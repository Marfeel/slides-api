package service;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;

import java.io.IOException;
import java.util.List;

import static java.util.Arrays.asList;
import static java.util.Collections.singletonList;

public class SheetsService {

    private final Sheets service;

    public SheetsService(Sheets service) {
        this.service = service;
    }

    public BatchUpdateSpreadsheetResponse batchUpdate(Spreadsheet spreadsheet, List<Request> sheetRequests) throws IOException {
        BatchUpdateSpreadsheetRequest sheetsReq = new BatchUpdateSpreadsheetRequest().setRequests(sheetRequests);
        return service.spreadsheets().batchUpdate(spreadsheet.getSpreadsheetId(), sheetsReq).execute();
    }

    public Spreadsheet createSpreadsheet(String mediaGroupName) throws IOException {
        Spreadsheet spreadsheet = service.spreadsheets()
                .create(new Spreadsheet()
                        .setProperties(new SpreadsheetProperties()
                                .setTitle(mediaGroupName))
                        .setSheets(asList(
                                new Sheet()
                                        .setProperties(new SheetProperties()
                                                .setTitle("Info for chart"))
                                        .setData(singletonList(
                                                new GridData()
                                                        .setStartColumn(0)
                                                        .setStartRow(0)
                                                        .setRowData(asList(
                                                                new RowData().setValues(asList( dateCellData("01/01/2017", "dd/mm/yyyy"), numberCellData(3.0)) ),
                                                                new RowData().setValues(asList( dateCellData("02/01/2017", "dd/mm/yyyy"), numberCellData(5.0)) ),
                                                                new RowData().setValues(asList( dateCellData("03/01/2017", "dd/mm/yyyy"), numberCellData(1.0)) ),
                                                                new RowData().setValues(asList( dateCellData("04/01/2017", "dd/mm/yyyy"), numberCellData(4.0))) )
                                                        )
                                        )),
                                new Sheet()
                                        .setProperties(new SheetProperties()
                                                .setTitle("Metrics Summary"))
                                        .setData(singletonList(
                                                new GridData()
                                                        .setStartColumn(0)
                                                        .setStartRow(0)
                                                        .setRowData(asList(
                                                                new RowData().setValues(asList(
                                                                        stringCellData(""),
                                                                        stringCellData("9 Feb '17  - 14 Feb '17"),
                                                                        stringCellData("16 Feb '17  - 27 Feb '17"),
                                                                        stringCellData("1 Mar '17  - 12 Mar '17"),
                                                                        stringCellData(""),
                                                                        stringCellData(""))
                                                                ),
                                                                new RowData().setValues(asList(
                                                                        stringCellData(""),
                                                                        stringCellData("B1 - Marfeel Touch"),
                                                                        stringCellData("B2 - Marfeel K"),
                                                                        stringCellData("B3 - Marfeel Touch"),
                                                                        stringCellData("Var. B2-B1 %"),
                                                                        stringCellData("Var. B3-B1 %"))
                                                                ),
                                                                new RowData().setValues(asList(
                                                                        stringCellData("ARPU"),
                                                                        formattedCellData(0.86, "CURRENCY"),
                                                                        formattedCellData(0.66, "CURRENCY"),
                                                                        formattedCellData(0.95, "CURRENCY"),
                                                                        formulaCellData("=C3/B3 - 1", "PERCENT"),
                                                                        formulaCellData("=D3/B3 - 1", "PERCENT"))
                                                                ),
                                                                new RowData().setValues(asList(
                                                                        stringCellData("Impression/V"),
                                                                        numberCellData(7.04),
                                                                        numberCellData(4.50),
                                                                        numberCellData(6.47),
                                                                        formulaCellData("=C4/B4 - 1", "PERCENT"),
                                                                        formulaCellData("=D4/B4 - 1", "PERCENT")))
                                                        ))
                                        ))
                        ))
                )
                .execute();

        return spreadsheet;
    }
    
    private CellData stringCellData(String text) {
        return new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(text));
    }

    private CellData numberCellData(Double number) {
        return new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue(number));
    }

    private CellData formattedCellData(Double value, String type) {
        return new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue(value))
                .setUserEnteredFormat(new CellFormat().setNumberFormat(new NumberFormat().setType(type)));
    }

    private CellData formulaCellData(String formula, String type) {
        return new CellData().setUserEnteredValue(new ExtendedValue().setFormulaValue(formula))
                .setUserEnteredFormat(new CellFormat().setNumberFormat(new NumberFormat().setType(type).setPattern("00.00%")));
    }

    private CellData dateCellData(String date, String pattern) {
        return new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(date))
                .setUserEnteredFormat(new CellFormat().setNumberFormat(new NumberFormat().setType("DATE").setPattern(pattern)));
    }

    public Request generateEmbeddedChart(Spreadsheet spreadsheet) {
        return new com.google.api.services.sheets.v4.model.Request()
                .setAddChart(new AddChartRequest()
                        .setChart(new EmbeddedChart().setPosition(new EmbeddedObjectPosition().setSheetId(0))
                                .setSpec(new ChartSpec()
                                        .setTitle("ecpm Mensual")
                                        .setBasicChart(new BasicChartSpec()
                                                .setChartType("LINE")
                                                .setLegendPosition("BOTTOM_LEGEND")
                                                .setAxis(asList(
                                                        new BasicChartAxis().setPosition("BOTTOM_AXIS").setTitle("Date"),
                                                        new BasicChartAxis().setPosition("LEFT_AXIS").setTitle("eCPM")
                                                ))
                                                .setDomains(singletonList(
                                                        new BasicChartDomain()
                                                                .setDomain(new ChartData()
                                                                        .setSourceRange(new ChartSourceRange()
                                                                                .setSources(singletonList(
                                                                                        new GridRange() //  [startIndex, endIndex)
                                                                                                .setSheetId(spreadsheet.getSheets().get(0).getProperties().getSheetId())
                                                                                                .setStartColumnIndex(0)
                                                                                                .setStartRowIndex(0)
                                                                                                .setEndColumnIndex(1)
                                                                                                .setEndRowIndex(4)
                                                                                ))
                                                                        ))
                                                ))
                                                .setSeries(singletonList(
                                                        new BasicChartSeries()
                                                                .setTargetAxis("LEFT_AXIS")
                                                                .setSeries(new ChartData().setSourceRange(new ChartSourceRange()
                                                                        .setSources(singletonList(
                                                                                new GridRange() //  [startIndex, endIndex)
                                                                                        .setSheetId(spreadsheet.getSheets().get(0).getProperties().getSheetId())
                                                                                        .setStartColumnIndex(1)
                                                                                        .setStartRowIndex(0)
                                                                                        .setEndColumnIndex(2)
                                                                                        .setEndRowIndex(4)
                                                                        ))))))
                                        )
                                ))
                );
    }

    public List<Request> conditionalFormatting(Spreadsheet spreadsheet) {
        Integer sheetId = spreadsheet.getSheets().stream().filter(sheet -> sheet.getProperties().getTitle().equals("Metrics Summary")).findFirst().get().getProperties().getSheetId();
        return asList(new Request()
                .setAddConditionalFormatRule(new AddConditionalFormatRuleRequest()
                        .setRule(new ConditionalFormatRule()
                                .setRanges(singletonList(new GridRange()
                                        .setSheetId(sheetId)
                                        .setStartRowIndex(2)
                                        .setEndRowIndex(4)
                                        .setStartColumnIndex(4)
                                        .setEndColumnIndex(6)))
                                .setBooleanRule(new BooleanRule()
                                        .setCondition(new BooleanCondition()
                                                .setType("NUMBER_LESS")
                                                .setValues(singletonList(new ConditionValue()
                                                        .setUserEnteredValue("-0.10"))))
                                        .setFormat(new CellFormat()
                                                .setTextFormat(new TextFormat()
                                                        .setForegroundColor(new Color()
                                                                .setRed(1f)))
                                        )
                                )
                        )
                ), new Request()
                .setAddConditionalFormatRule(new AddConditionalFormatRuleRequest()
                        .setRule(new ConditionalFormatRule()
                                .setRanges(singletonList(new GridRange()
                                        .setSheetId(sheetId)
                                        .setStartRowIndex(2)
                                        .setEndRowIndex(4)
                                        .setStartColumnIndex(4)
                                        .setEndColumnIndex(6)))
                                .setBooleanRule(new BooleanRule()
                                        .setCondition(new BooleanCondition()
                                                .setType("NUMBER_GREATER")
                                                .setValues(singletonList(new ConditionValue()
                                                        .setUserEnteredValue("0.10"))))
                                        .setFormat(new CellFormat()
                                                .setTextFormat(new TextFormat()
                                                        .setForegroundColor(new Color()
                                                                .setRed(0.133333f)
                                                                .setGreen(0.545098f)
                                                                .setBlue(0.133333f)))
                                        )
                                )
                        )
                ), new Request()
                .setAddConditionalFormatRule(new AddConditionalFormatRuleRequest()
                        .setRule(new ConditionalFormatRule()
                                .setRanges(singletonList(new GridRange()
                                        .setSheetId(sheetId)
                                        .setStartRowIndex(2)
                                        .setEndRowIndex(4)
                                        .setStartColumnIndex(4)
                                        .setEndColumnIndex(6)))
                                .setBooleanRule(new BooleanRule()
                                        .setCondition(new BooleanCondition()
                                                .setType("NUMBER_BETWEEN")
                                                .setValues(asList(
                                                        new ConditionValue()
                                                                .setUserEnteredValue("-0.10"),
                                                        new ConditionValue()
                                                                .setUserEnteredValue("0.10"))))
                                        .setFormat(new CellFormat()
                                                .setTextFormat(new TextFormat()
                                                        .setForegroundColor(new Color()
                                                                .setRed(0f)
                                                                .setGreen(0f)
                                                                .setBlue(1f)))
                                        )
                                )
                        )
                ));

    }

}
