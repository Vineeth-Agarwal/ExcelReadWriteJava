/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package read_write;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

/**
 *
 * @author S530671
 */
public class ReadWrite {

    /**
     * @param args the command line arguments
     * @throws java.io.FileNotFoundException
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException
     * @throws java.text.ParseException
     */
    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException, ParseException {
        // TODO code application logic here

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Output");

        //Name Styling and content
        Row nameRow = sheet.createRow(1);
        CellStyle nameStyle;
        Font font;

        Cell nameCell = nameRow.createCell(1);
        nameCell.setCellValue("Name");
        nameStyle = workbook.createCellStyle();
        font = workbook.createFont();
        font.setBold(true);
        nameStyle.setFont(font);
        nameStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        nameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        nameStyle.setBorderBottom(BorderStyle.THIN);
        nameStyle.setBorderTop(BorderStyle.THIN);
        nameStyle.setBorderLeft(BorderStyle.THIN);
        nameStyle.setBorderRight(BorderStyle.THIN);
        nameCell.setCellStyle(nameStyle);

        Cell fullNamecell = nameRow.createCell(2);
        fullNamecell.setCellValue("Agarwal, Vineeth");
        nameStyle = workbook.createCellStyle();
        font = workbook.createFont();
        font.setItalic(true);
        font.setUnderline(Font.U_SINGLE);
        nameStyle.setFont(font);
        nameStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        nameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        nameStyle.setBorderBottom(BorderStyle.THIN);
        nameStyle.setBorderTop(BorderStyle.THIN);
        nameStyle.setBorderLeft(BorderStyle.THIN);
        nameStyle.setBorderRight(BorderStyle.THIN);
        fullNamecell.setCellStyle(nameStyle);

        Cell MergeName = nameRow.createCell(3);
        MergeName.setCellStyle(nameStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 2, 3));

        //Header Styling and content
        Row headerRow = sheet.createRow(3);
        CellStyle headerCellStyle = workbook.createCellStyle();
        font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        headerCellStyle.setFont(font);
        headerCellStyle.setFillForegroundColor(IndexedColors.fromInt(2).getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        headerCellStyle.setBorderBottom(BorderStyle.THIN);
        headerCellStyle.setBorderLeft(BorderStyle.THIN);
        headerCellStyle.setBorderRight(BorderStyle.THIN);
        headerCellStyle.setBorderTop(BorderStyle.THIN);

        Cell cellhead0 = headerRow.createCell(0);
        cellhead0.setCellValue("SNO");
        cellhead0.setCellStyle(headerCellStyle);
        Cell cellhead1 = headerRow.createCell(1);
        cellhead1.setCellValue("Genre");
        cellhead1.setCellStyle(headerCellStyle);
        Cell cellhead2 = headerRow.createCell(2);
        cellhead2.setCellValue("Critic Score");
        cellhead2.setCellStyle(headerCellStyle);
        Cell cellhead3 = headerRow.createCell(3);
        cellhead3.setCellValue("Album Name");
        cellhead3.setCellStyle(headerCellStyle);
        Cell cellhead4 = headerRow.createCell(4);
        cellhead4.setCellValue("Artist");
        cellhead4.setCellStyle(headerCellStyle);
        Cell cellhead5 = headerRow.createCell(5);
        cellhead5.setCellValue("Release Date");
        cellhead5.setCellStyle(headerCellStyle);

        int count = 4;
        TreeMap<String, ArrayList<String>> tm1 = new TreeMap<>();
        ArrayList<String> values;
        File inputFile = new File("Agarwal_Input.xlsx");
        XSSFWorkbook input_workbook = new XSSFWorkbook(inputFile);
        XSSFSheet input_sheet = input_workbook.getSheetAt(0);
        for (int i = 1; i < input_sheet.getPhysicalNumberOfRows(); i++) {
            Row input_row = input_sheet.getRow(i);
            String input[] = {"" + (int) input_row.getCell(0).getNumericCellValue(),
                input_row.getCell(1).getStringCellValue(), input_row.getCell(2).getStringCellValue(),
                input_row.getCell(3).getStringCellValue(), "" + input_row.getCell(4).getDateCellValue(),
                "" + (int) input_row.getCell(5).getNumericCellValue()
            };

            if (tm1.containsKey(input[2])) {
                tm1.get(input[2]).add(input[0] + "\t" + input[1] + "\t" + input[2] + "\t" + input[3] + "\t" + input[4]
                        + "\t" + input[5]);
            } else {
                values = new ArrayList<>();
                values.add(input[0] + "\t" + input[1] + "\t" + input[2] + "\t" + input[3] + "\t" + input[4] + "\t" + input[5]);
                tm1.put(input[2], values);
            }

        }

        FileOutputStream f_out = new FileOutputStream(new File("output_workbook.xlsx"));
        String genre = "";
        ArrayList<Album> albumList = new ArrayList<>();

        for (ArrayList<String> aList : tm1.values()) {

            for (String str : aList) {
                String[] b = str.trim().split("\t");

                Album a1 = new Album(Integer.parseInt(b[0]), b[2], Integer.parseInt(b[5]), b[1], b[3], b[4]);
                albumList.add(a1);
            }
        }
//        Collections.sort(albumAL, new AlbumSort());
        Random rand = new Random();
        float r = rand.nextFloat();
        float g = rand.nextFloat();
        float b = rand.nextFloat();
        XSSFColor color = new XSSFColor(new java.awt.Color(r, g, b));

        for (Album album : albumList) {
            album.SNO = ++count - 4;
            Row row = sheet.createRow(count - 1);
            CellStyle tan_style = workbook.createCellStyle();

            tan_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            tan_style.setBorderBottom(BorderStyle.THIN);
            tan_style.setBorderLeft(BorderStyle.THIN);
            tan_style.setBorderRight(BorderStyle.THIN);
            tan_style.setBorderTop(BorderStyle.THIN);

            if (!genre.equals(album.genre)) {
                r = rand.nextFloat() / 2f + (float) 0.5;
                g = rand.nextFloat() / 2f + (float) 0.5;
                b = rand.nextFloat() / 2f + (float) 0.5;
                XSSFColor fillColor = new XSSFColor(new java.awt.Color(r, g, b));
                ((XSSFCellStyle) tan_style).setFillForegroundColor(fillColor);
                color = fillColor;
            } else {
                ((XSSFCellStyle) tan_style).setFillForegroundColor(color);
            }

            genre = album.genre;

            Cell cellSno = row.createCell(0);
            Cell cellGenre = row.createCell(1);
            Cell cellCriticScore = row.createCell(2);
            Cell cellAlbumName = row.createCell(3);
            Cell cellArtist = row.createCell(4);
            Cell cellReleaseDate = row.createCell(5);

            cellSno.setCellValue(album.SNO);
            cellGenre.setCellValue(album.genre);
            cellCriticScore.setCellValue(album.criticScore);
            cellAlbumName.setCellValue(album.albumName);
            cellArtist.setCellValue(album.artist);
            SimpleDateFormat parseFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy");
            Date date = parseFormat.parse(album.date);
            SimpleDateFormat format = new SimpleDateFormat("dd-MMM-yy");
            String result = format.format(date);
            cellReleaseDate.setCellValue(result);

            cellSno.setCellStyle(tan_style);
            cellGenre.setCellStyle(tan_style);
            cellCriticScore.setCellStyle(tan_style);
            cellAlbumName.setCellStyle(tan_style);
            cellArtist.setCellStyle(tan_style);
            cellReleaseDate.setCellStyle(tan_style);
        }

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);
        sheet.autoSizeColumn(5);
        workbook.write(f_out);
        f_out.close();
    }
}
