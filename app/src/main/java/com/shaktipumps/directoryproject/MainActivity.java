package com.shaktipumps.directoryproject;

import android.Manifest;
import android.app.Activity;
import android.content.Context;
import android.content.pm.PackageManager;
import android.os.Environment;
import android.os.Bundle;
import android.util.Log;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    Context context;
    public String dirName = "";
    public  static   boolean success = false;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        context = this;

        storeExcelFile();
    }



    private void storeExcelFile() {
        if (ActivityCompat.checkSelfPermission(context, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                != PackageManager.PERMISSION_GRANTED &&ActivityCompat.checkSelfPermission(context, Manifest.permission.READ_MEDIA_IMAGES)
                != PackageManager.PERMISSION_GRANTED) {

            ActivityCompat.requestPermissions((Activity) context,
                    new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE,Manifest.permission.READ_MEDIA_IMAGES}, 1);

            return;
        }
        dirName = getMediaFilePath("testing","Anjali.xls");
        if(!dirName.isEmpty()){
            saveExcelFile(dirName);
        }
    }

    public String getMediaFilePath(String type,String name) {

        File root = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS), "GALLERY_DIRECTORY_NAME_COMMON");

        File dir = new File(root.getAbsolutePath() + "/SKAPP/"+ type ); //it is my root directory

        try {
            if (!dir.exists()) {
                dir.mkdirs();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        // Create a media file name
        return dir.getPath() + File.separator + name;
    }

    private  void saveExcelFile(String dirName) {

        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.e("Failed", "Storage not available or read only");
        }
         //New Workbook
        Workbook wb = new HSSFWorkbook();

        Cell c = null;

        //Cell style for header row

        Font font = wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);

        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("myOrder");

        // Generate column headings
        Row row = null;

        row = sheet1.createRow(0);

        c = row.createCell(0);
        c.setCellValue("Item Number");


        c = row.createCell(1);
        c.setCellValue("Quantity");

        c = row.createCell(2);
        c.setCellValue("Price");

        sheet1.setColumnWidth(0, (15 * 500));
        sheet1.setColumnWidth(1, (15 * 500));
        sheet1.setColumnWidth(2, (15 * 500));

        int val = 0;
        int k = 1;
        for (int i = 1; i < 12; i++) {
            row = sheet1.createRow(k);
            for (int j = 0; j < 3; j++) {
                c = row.createCell(j);
                c.setCellValue(val);

                val++;
            }
            sheet1.setColumnWidth(i, (15 * 500));
            k++;
        }


        try {
            wb.createSheet("Anjali.xls");
            FileOutputStream os = new FileOutputStream(dirName);

            wb.write(os);
            os.close();
            Log.w("FileUtils", "Writing file" + "Sample.xls");
            success = true;
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " +"Sample.xls", e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);

        }

    }

    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) {
            return true;
        }
        return false;
    }
}