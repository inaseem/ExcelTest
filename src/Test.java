
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author Naseem
 */
public class Test {

    private static int rowIndex = 0;

    public static void main(String args[]) {
        //New Workbook
        Workbook wb = new HSSFWorkbook();

        Cell c = null;

        //Cell style for header row
        CellStyle cs = wb.createCellStyle();
        cs.setFillForegroundColor(HSSFColor.LIME.index);

        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("myOrder");

        // Generate column headings
        Row row = sheet1.createRow(0);

        c = row.createCell(0);
        c.setCellValue("Item Number\n");
        c.setCellStyle(cs);

        c = row.createCell(1);
        c.setCellValue("Quantity");
        c.setCellStyle(cs);

        c = row.createCell(2);
        c.setCellValue("Price");
        c.setCellStyle(cs);

        sheet1.setColumnWidth(0, (15 * 500));
        sheet1.setColumnWidth(1, (15 * 500));
        sheet1.setColumnWidth(2, (15 * 500));
        sheet1.autoSizeColumn(0); //adjust width of the first column
        sheet1.autoSizeColumn(1);
        for (int i = 0; i < 10; ++i) {
            sheet1.createRow(++rowIndex).createCell(0).setCellValue("Naseem");
        }
        CellStyle style;
        Font titleFont = wb.createFont();
        //titleFont.setFontHeightInPoints((short)18);
        //titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
        titleFont.setBold(true);
        style = wb.createCellStyle();
        //style.setShrinkToFit(true);
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        Row headerRow = sheet1.createRow(++rowIndex);
        headerRow.setHeightInPoints(60);
        Cell titleCell = headerRow.createCell(0);
        titleCell.setCellValue("Head\nHead\nHead\nHead");
        titleCell.setCellStyle(style);
        sheet1.addMergedRegion(CellRangeAddress.valueOf("$A$" + (rowIndex + 1) + ":$N$" + (rowIndex + 1)));

        FileOutputStream os = null;
        try {
            File file=new File("naseem.jpg");
            InputStream is = new FileInputStream(file);
            byte[] bytes = IOUtils.toByteArray(is);
            int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            is.close();
            CreationHelper helper = wb.getCreationHelper();
            // Create the drawing patriarch.  This is the top level container for all shapes. 
            Drawing drawing = sheet1.createDrawingPatriarch();

            //add a picture shape
            ClientAnchor anchor = helper.createClientAnchor();
            //set top-left corner of the picture,
            //subsequent call of Picture#resize() will operate relative to it
            anchor.setCol1(0);
            anchor.setRow1(++rowIndex);
            //anchor.setCol2(10);
            //anchor.setRow2(5);
            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE);
            Picture pict = drawing.createPicture(anchor, pictureIdx);

            //auto-size picture relative to its top-left corner
            pict.resize(5, 5);
            os = new FileOutputStream(new File("naseem.xls"));
            wb.write(os);
            //Log.w("FileUtils", "Writing file" + file); 
            //success = true; 
            File file2=new File("/");
            System.out.println(file2.getAbsolutePath());
        } catch (IOException e) {
            System.out.println(e.getMessage());
            //Log.w("FileUtils", "Error writing " + file, e); 
        } catch (Exception e) {
            System.out.println(e.getMessage());
            //Log.w("FileUtils", "Failed to save file", e); 
        } finally {
            try {
                if (null != os) {
                    os.close();
                }
            } catch (Exception ex) {
            }
        }
    }
}
