package com.example.applicationexcel;

import static android.content.Intent.FLAG_GRANT_READ_URI_PERMISSION;
import static android.content.Intent.FLAG_GRANT_WRITE_URI_PERMISSION;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.documentfile.provider.DocumentFile;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

public class MainActivity extends AppCompatActivity {

    Button button;
    TextView slx;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        //  XSSFWorkbook workbook = new XSSFWorkbook();
        button = findViewById(R.id.button);
        slx = findViewById(R.id.slx);
        button.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                callIntent();
            }
        });
    }

    public void callIntent() {
        Intent intent = new Intent()
                .setType("*/*")
                .setAction(Intent.ACTION_GET_CONTENT);
        intent.putExtra(Intent.EXTRA_LOCAL_ONLY, true);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setFlags(FLAG_GRANT_READ_URI_PERMISSION | FLAG_GRANT_WRITE_URI_PERMISSION);
        startActivityForResult(Intent.createChooser(intent, "Select a file"), 123);
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == 123 && resultCode == RESULT_OK) {
            //getExcelFileFromStorage(new onGetExcelFill(){
            //    @Override
            //    public void onGetFile(File file) {
            //        String [][] readFile = read(file);
            //        Toast.makeText(MainActivity.this, "try again", Toast.LENGTH_SHORT).show();

            //    }
            //});

            XSSFWorkbook wb = null;
            HSSFWorkbook wb2 = null;
            if ((data != null) && (data.getData() != null)) {
                Uri selectedFile = data.getData();
                try {    InputStream inputStream
                        = this.getContentResolver().openInputStream(selectedFile);

                    POIFSFileSystem poifsFileSystem=new POIFSFileSystem(inputStream);
                    wb2=new HSSFWorkbook(poifsFileSystem);
                    Sheet sheet12 = wb2.getSheetAt(0);

                    Log.d("dpflfdkfl", sheet12.getRow(6).getCell(3).toString());
                    Log.d("dpflfdkfl", "dd");
                  String extension =
                         DocumentFile.fromSingleUri(this, selectedFile).getName();
                    Log.d("dpflfsdkfl", extension.substring(extension.lastIndexOf(".") + 1));




                    //.name?.takeLastWhile { char ->
                      //  char != '.' }



       //          wb = new XSSFWorkbook(inputStream);

       //          Sheet sheet1 = wb.getSheetAt(0);
       //        //  sheet1.getRow(1).getCell(1).toString();
       //        //  sheet1.getColumnOutlineLevel(1);
       //        //  sheet1.getActiveCell();
       //    //     Log.d("dpflfdkfl", sheet1.getRow(1).getCell(1).toString());
       //    //     Log.d("dpflfdkfl", sheet1.getRow(1).getCell(2).toString());
       //    //     Log.d("dpflfdkfl", sheet1.getRow(1).getCell(3).toString());
       //    //     Log.d("dpflfdkfl", sheet1.getRow(1).getCell(4).toString());
       //    //     Log.d("dpflfdkfl", sheet1.getRow(1).getCell(5).toString());
       //    //     Log.d("dpflfdkfl", sheet1.getRow(1).getCell(6).toString());
       //    //     Log.d("dpflfdkfl","...........");

       //    //     Log.d("dpflfdkfl", sheet1.getRow(2).getCell(1).toString());
       //    //     Log.d("dpflfdkfl", sheet1.getRow(2).getCell(2).toString());
       //    //     Log.d("dpflfdkfl", sheet1.getRow(2).getCell(3).toString());
       //    //     Log.d("dpflfdkfl", sheet1.getRow(2).getCell(4).toString());
       //          //
                    // if (GetFileExtension(FilePath).equals(".xls")) {

                    //     wb = new HSSFWorkbook(inStream);
                    //     sheet1 = wb.getSheetAt(0);
                    //     Formeval = new HSSFFormulaEvaluator((HSSFWorkbook) wb);

                    // }else if (GetFileExtension(FilePath).equals(".xlsx")) {

                    //     wb = new XSSFWorkbook(inStream);
                    //     sheet1 = wb.getSheetAt(0);
                    //     Formeval = new XSSFFormulaEvaluator((XSSFWorkbook) wb);
                    // }
                    //  InputStream inputStream
                    //          = this .getContentResolver().openInputStream(selectedFile);
                    //  // POIFSFileSystem s=
                    //XSSFWorkbook workbook = new XSSFWorkbook();
                    // workbook.


                    //  InputStream ExcelFileToRead = new FileInputStream("C:/Test.xlsx");
                    //  XSSFWorkbook  workbook = new XSSFWorkbook(ExcelFileToRead);

                    // XSSFSheet mySheet  = workbook.getSheetAt(0);
                    //    Log.d("dpflfdkfl", mySheet.getSheetName());


                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                    Log.d("dpflfdkfl", e.getMessage());

                } catch (IOException e) {
                    Log.d("dpflfdkfl", e.getMessage());

                    e.printStackTrace();
                }
                //contentResolver.openInputStream(uri)

                File file = new File(String.valueOf(selectedFile));

                String strFileName = file.getName();

                //   workbook = null;

                //  WorkbookSettings ws = new WorkbookSettings();
                //  ws.setGCDisabled(true);

                Toast.makeText(MainActivity.this, "try again63", Toast.LENGTH_SHORT).show();


                ///  Workbook     workbook = Workbook.getWorkbook(file, ws);
                ///      Sheet sheet = workbook.getSheet(0);
                ///      sheet.getColumn(0);
                ///  XSSFWorkbook workbook = new XSSFWorkbook(stream);
//                Workbook workbook = new HSSFWorkbook();
//                Sheet sheet = workbook.createSheet("My Sheet");
//                Keyboard.Row row= sheet.createRow(0);
//                Cell cell = row.createCell(0);
//
//                myWorkbook = new XSSFWorkbook(OPCPackage.open(file));

                Intent intent;

                String[] mimetypes =
                        {"application/vnd.ms-excel", // .xls
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" // .xlsx
                        };

                intent = new Intent(Intent.ACTION_GET_CONTENT); // or use ACTION_OPEN_DOCUMENT
                intent.setType("*/*");
                intent.putExtra(Intent.EXTRA_MIME_TYPES, mimetypes);
                intent.addCategory(Intent.CATEGORY_OPENABLE);


                // Toast.makeText(this,  sheet.getColumn(0).toString(), Toast.LENGTH_SHORT).show();
                Toast.makeText(this, selectedFile.getPath().toString(), Toast.LENGTH_SHORT).show();
                //  read(file);
                // InputStream stream = getContentResolver().openInputStream(selectedFile);
                // XSSFWorkbook workbook = new XSSFWorkbook(stream);
                // Toast.makeText(MainActivity.this, selectedFile.getPath().toString(), Toast.LENGTH_SHORT).show();

            }
        }
    }
    //  private static void readFile(Context context, String filename) {

    //      if (!isExternalStorageAvailable() || isExternalStorageReadOnly())
    //      {
    //          Log.w("FileUtils", "Storage not available or read only");
    //          return;
    //      }

    //      FileInputStream fis = null;

    //      try
    //      {
    //          File file = new File(context.getExternalFilesDir(null), filename);
    //          fis = new FileInputStream(file);
    //          // Get the object of DataInputStream
    //          DataInputStream in = new DataInputStream(fis);
    //          BufferedReader br = new BufferedReader(new InputStreamReader(in));
    //          String strLine;
    //          //Read File Line By Line
    //          while ((strLine = br.readLine()) != null) {
    //              Log.w("FileUtils", "File data: " + strLine);
    //              Toast.makeText(context, "File Data: " + strLine , Toast.LENGTH_SHORT).show();
    //          }
    //          in.close();
    //      }
    //      catch (Exception ex) {
    //          Log.e("FileUtils", "failed to load file", ex);
    //      }
    //      finally {
    //          try {if (null != fis) fis.close();} catch (IOException ex) {}
    //      }

    //      return;
    //  }

    public StringBuffer splitString(String str)
    {
        //  ArrayList CompanyId =new ArrayList<>();
        StringBuffer type = new StringBuffer();
        // StringBuffer    num = new StringBuffer();
        for (int i=0; i<str.length(); i++)
        {
            if(i>2)
            {
                type .append(str.charAt(i));
            }
        }
        return type;
    }
}