package com.mithu.excelDownload

import grails.transaction.Transactional
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Workbook
import pl.touk.excel.export.WebXlsxExporter

@Transactional
class ExcelDownloadService {
    def save(String fileNamePrefix, List headers, List properties, def data, def response, List totalFields = []) {
        Integer lastRowNum = data.size() + 1

        try {
            String fileName = fileNamePrefix  + '.xlsx'

            WebXlsxExporter webXlsxExporter = new WebXlsxExporter()
            setHeaderCellStyle(webXlsxExporter, headers)
            webXlsxExporter.with {
                setResponseHeaders(response, fileName)
                fillHeader(headers)
                add(data, properties)
            }
            if (data && totalFields) {
                addTotalRow(webXlsxExporter, headers, lastRowNum, totalFields)
                setTotalRowCellStyle(webXlsxExporter, headers, lastRowNum)
            }
            webXlsxExporter.save(response.outputStream)
        } catch (Exception e) {
            throw new Exception("Error in saving Excel File")
        }

    }

    private setHeaderCellStyle(WebXlsxExporter webXlsxExporter, List headers) {
        Workbook wb = webXlsxExporter.getWorkbook()
        headers.eachWithIndex() { item, i ->
            Font headerFont = wb.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 9);
            headerFont.setColor(IndexedColors.WHITE.getIndex())
            webXlsxExporter.putCellValue(0, i, item.value.toString())

            def cellStyle = webXlsxExporter.getCellAt(0, i).getCellStyle()
            cellStyle.setFont(headerFont);
            cellStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            webXlsxExporter.getCellAt(0, i).setCellStyle(cellStyle)
        }
    }

    private setTotalRowCellStyle(WebXlsxExporter webXlsxExporter, List headers, Integer lastRowNum) {
        headers?.eachWithIndex { def item, Integer i ->
            def cellStyle = webXlsxExporter.getCellAt(lastRowNum, i).getCellStyle()
            cellStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            webXlsxExporter.getCellAt(lastRowNum, i).setCellStyle(cellStyle)
        }
    }


    private addTotalRow(WebXlsxExporter webXlsxExporter, List headers, Integer lastRowNum, List totalFields) {
        List totalRow = []
        Integer columnNos = headers.size()
        if (columnNos > 0) {
            totalRow.add("Total")
            //adding a last empty row
            (2..columnNos).eachWithIndex { def colValue, int index ->
                totalRow.add("")
            }
            webXlsxExporter.with {
                fillRow(totalRow, lastRowNum)
            }

            totalFields?.each { item ->
                Integer index = headers.findIndexOf { it == item }
                ('A'..'Z').eachWithIndex { String entry, int i ->
                    if (i == index) {
                        String ref = entry + 2 + ":" + entry + (lastRowNum)
                        webXlsxExporter.getCellAt(lastRowNum, i).setCellFormula("SUM(" + ref + ")");
                    }
                }
            }
        }

    }
}
