
import ali.Data;
import ali.Item;
import com.google.gson.Gson;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author Naseem
 */
public class Reader {

    public static void main(String[] args) throws Exception {
        OkHttpClient client = new OkHttpClient();
        FileInputStream fis = new FileInputStream("input.xls");
        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(fis));
        Workbook outbook = new HSSFWorkbook();
        Sheet outsheet = outbook.createSheet();
        Sheet sheet = wb.getSheetAt(0);
        final Gson gson = new Gson();
        int num = sheet.getLastRowNum();
        for (int i = 0; i < num; ++i) {
            Row row = sheet.getRow(i);
            Row outrow = outsheet.createRow(i);
            int index = 0;
            Iterator<Cell> it = row.cellIterator();
            if (i > 0) {
                String ISBN = "";
                while (it.hasNext()) {
                    Cell cell = it.next();
                    Cell c = outrow.createCell(index++);
                    if (index == 1) {
                        ISBN = String.format("%.0f", cell.getNumericCellValue());
                        System.out.print(i+". "+ISBN);
                        c.setCellValue(ISBN);
                    }
                    if (index == 2 && ISBN != null) {
                        Request request = new Request.Builder()
                                .url("https://www.googleapis.com/books/v1/volumes?q=" + ISBN)
                                .build();

                        Response response = client.newCall(request).execute();
                        String res = response.body().string();
                        if (res != null) {
                            Data data = gson.fromJson(res, Data.class);
                            if (data.getItems() != null) {
                                if (data.getItems().size() > 0) {
                                    Item item = data.getItems().get(0);
                                    c.setCellValue(item.getVolumeInfo().getTitle());
                                    System.out.println(" = "+item.getVolumeInfo().getTitle());
                                }
                            }
                        }
                    }
                }

//                System.out.println("Done " + ISBN);
            } else {
                while (it.hasNext()) {
                    Cell cell = it.next();
                    Cell c = outrow.createCell(index++);
                    c.setCellValue(cell.getStringCellValue());
                }
            }
        }
        FileOutputStream fos = new FileOutputStream("Box-24_out.xls");
        outbook.write(fos);
    }
}
