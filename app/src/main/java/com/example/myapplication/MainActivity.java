package com.example.myapplication;

import android.Manifest;
import android.content.ContentUris;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.DocumentsContract;
import android.provider.MediaStore;
import android.provider.Settings;
import android.telephony.SmsManager;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

import com.google.android.material.floatingactionbutton.FloatingActionButton;
import com.google.android.material.textfield.TextInputLayout;

import org.apache.poi.hssf.record.PageBreakRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;


public class MainActivity extends AppCompatActivity {

    private TextInputLayout message, number;

    private Button send, choosefile;

    private Intent myFileIntent;

    private TextView textview;

    private FloatingActionButton floatingbutton2,floatingbutton1;

    private static final String TAG = "MainActivity";

    String FilePath;
    MainActivity c;

    int cellnumber;
    int rownumber;

    int count = 0;
    File file;

   ArrayList<String> MobNo = new ArrayList<String>();
   ArrayList<String> DelMessage = new ArrayList<String>();


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        number = (TextInputLayout) findViewById(R.id.MobileNumber);
        message = (TextInputLayout) findViewById(R.id.Message);
        send = (Button) findViewById(R.id.button);
        choosefile = (Button) findViewById(R.id.ChooseFile);
        textview = (TextView) findViewById(R.id.textView);
        floatingbutton2 = (FloatingActionButton) findViewById(R.id.floatingActionButton2);
        floatingbutton1 = (FloatingActionButton) findViewById(R.id.floatingActionButton);
        floatingbutton1.setEnabled(false);

        send.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                //check condition
                if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.SEND_SMS) == PackageManager.PERMISSION_GRANTED) {
                    sendSMS();
                } else {
                    ActivityCompat.requestPermissions(MainActivity.this, new String[]
                            {Manifest.permission.SEND_SMS}, 100);
                }
            }
        });


        choosefile.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {

                //check permission
                if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.READ_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED) {

                    if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED) {
                        FileChooser();
                    }
                } else {
                    ActivityCompat.requestPermissions(MainActivity.this, new String[]
                            {Manifest.permission.READ_EXTERNAL_STORAGE}, 101);

                    ActivityCompat.requestPermissions(MainActivity.this, new String[]
                            {Manifest.permission.WRITE_EXTERNAL_STORAGE}, 102);
                }

            }
        });


        floatingbutton2.setOnClickListener((new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if(FilePath!=null) {

                    if (MobNo.size() != 0) {
                        floatingbutton1.setEnabled(true);
                        System.out.println(MobNo);
                        System.out.println(DelMessage);
                        if (count < MobNo.size()) {
                            //System.out.println(MobNo.get(count));
                            number.getEditText().setText(MobNo.get(count));
                            //System.out.println(DelMessage.get(count));
                            message.getEditText().setText(DelMessage.get(count));
                            count++;
                        } else if (count == MobNo.size()) {
                            floatingbutton2.setEnabled(false);
                        }
                    }
                    else{
                        Toast.makeText(getApplicationContext(), "Please select a valid file", Toast.LENGTH_SHORT).show();
                    }
                }
                else{
                    Toast.makeText(getApplicationContext(), "Please select a valid file", Toast.LENGTH_SHORT).show();
                }

            }
        }));

        floatingbutton1.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if (FilePath!=null && MobNo.size()!=0) {
                    if (count > 0) {
                        floatingbutton2.setEnabled(true);
                        if (MobNo != null) {
                            //System.out.println(MobNo.get(count));

                            //System.out.println(DelMessage.get(count));
                            count--;
                            number.getEditText().setText(MobNo.get(count));
                            message.getEditText().setText(DelMessage.get(count));
                        }
                        else if(count == 0){
                           floatingbutton1.setEnabled(false);
                        }
                    }
                    else{
                        Toast.makeText(getApplicationContext(), "Please select a valid file", Toast.LENGTH_SHORT).show();
                    }
                } else{
                        floatingbutton1.setEnabled(false);
                    }

            }
        });

    }


    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        switch (requestCode) {
            case 10:
                if (resultCode == RESULT_OK) {
                    // FilePath = data.getData().getPath();
                    //File FilePath = Environment.getExternalStorageDirectory();
                    //File path = Environment.getExternalStoragePublicDirectory(Environment.getExternalStorageState());
                    Uri u = data.getData();
                    FilePath = c.getActualPath(this, u);
                    textview.setText(FilePath);
                    try {
                        ReadExcel();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }

                } else {
                    break;
                }
        }
    }

    public static String getActualPath(final Context context, final Uri uri) {

        final boolean isKitKat = Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT;

        // DocumentProvider
        if (isKitKat && DocumentsContract.isDocumentUri(context, uri)) {
            // ExternalStorageProvider
            if (isExternalStorageDocument(uri)) {
                final String docId = DocumentsContract.getDocumentId(uri);
                final String[] split = docId.split(":");
                final String type = split[0];

                if ("primary".equalsIgnoreCase(type)) {
                    return Environment.getExternalStorageDirectory() + "/" + split[1];
                }

                // TODO handle non-primary volumes
            }
            // DownloadsProvider
            else if (isDownloadsDocument(uri)) {

                final String id = DocumentsContract.getDocumentId(uri);
                final Uri contentUri = ContentUris.withAppendedId(
                        Uri.parse("content://downloads/public_downloads"), Long.valueOf(id));

                return getDataColumn(context, contentUri, null, null);
            }
            // MediaProvider
            else if (isMediaDocument(uri)) {
                final String docId = DocumentsContract.getDocumentId(uri);
                final String[] split = docId.split(":");
                final String type = split[0];

                Uri contentUri = null;
                if ("image".equals(type)) {
                    contentUri = MediaStore.Images.Media.EXTERNAL_CONTENT_URI;
                } else if ("video".equals(type)) {
                    contentUri = MediaStore.Video.Media.EXTERNAL_CONTENT_URI;
                } else if ("audio".equals(type)) {
                    contentUri = MediaStore.Audio.Media.EXTERNAL_CONTENT_URI;
                }

                final String selection = "_id=?";
                final String[] selectionArgs = new String[]{
                        split[1]
                };

                return getDataColumn(context, contentUri, selection, selectionArgs);
            }
        }
        // MediaStore (and general)
        else if ("content".equalsIgnoreCase(uri.getScheme())) {
            return getDataColumn(context, uri, null, null);
        }
        // File
        else if ("file".equalsIgnoreCase(uri.getScheme())) {
            return uri.getPath();
        }

        return null;
    }

    /**
     * Get the value of the data column for this Uri. This is useful for
     * MediaStore Uris, and other file-based ContentProviders.
     *
     * @param context       The context.
     * @param uri           The Uri to query.
     * @param selection     (Optional) Filter used in the query.
     * @param selectionArgs (Optional) Selection arguments used in the query.
     * @return The value of the _data column, which is typically a file path.
     */
    private static String getDataColumn(Context context, Uri uri, String selection,
                                        String[] selectionArgs) {

        Cursor cursor = null;
        final String column = "_data";
        final String[] projection = {
                column
        };

        try {
            cursor = context.getContentResolver().query(uri, projection, selection, selectionArgs,
                    null);
            if (cursor != null && cursor.moveToFirst()) {
                final int column_index = cursor.getColumnIndexOrThrow(column);
                return cursor.getString(column_index);
            }
        } finally {
            if (cursor != null)
                cursor.close();
        }
        return null;
    }


    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is ExternalStorageProvider.
     */
    public static boolean isExternalStorageDocument(Uri uri) {
        return "com.android.externalstorage.documents".equals(uri.getAuthority());
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is DownloadsProvider.
     */
    public static boolean isDownloadsDocument(Uri uri) {
        return "com.android.providers.downloads.documents".equals(uri.getAuthority());
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is MediaProvider.
     */
    public static boolean isMediaDocument(Uri uri) {
        return "com.android.providers.media.documents".equals(uri.getAuthority());
    }


    private void FileChooser() {
        myFileIntent = new Intent(Intent.ACTION_GET_CONTENT);
        myFileIntent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        //myFileIntent.addCategory(Intent.CATEGORY_OPENABLE);
        myFileIntent.setType("*/*");
        //myFileIntent.setType("application/vnd.ms-excel");
        String[] mimetypes = {"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"};
        myFileIntent.putExtra(myFileIntent.EXTRA_MIME_TYPES, mimetypes);
        startActivityForResult(myFileIntent, 10);
    }


    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (requestCode == 100 && grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
            //when permission is granted
            sendSMS();
        } else if (requestCode == 101 && grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
            //when permission is granted
            FileChooser();
        } else if (requestCode == 102 && grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {

            //when permission is granted
            FileChooser();
        } else {

            Toast.makeText(getApplicationContext(), "Permission_Denied!", Toast.LENGTH_SHORT).show();
        }
    }

    private void sendSMS() {
        //number = (EditText) findViewById(R.id.MobileNumber);
        //TextInputLayout text =findViewById(R.id.MobileNumber);
        String PhoneNo = number.getEditText().getText().toString().trim();

        //message = (EditText) findViewById(R.id.message)
        //TextInputLayout Mess = findViewById(R.id.message);
        String SMS = message.getEditText().getText().toString().trim();

        if (!PhoneNo.equals("") && !SMS.equals("")) {


            SmsManager smsManager = SmsManager.getDefault();
            smsManager.sendTextMessage(PhoneNo, null, SMS, null, null);
            Toast.makeText(getApplicationContext(), "Sms sent successfully", Toast.LENGTH_LONG).show();
        } else {
            //value is blank
            Toast.makeText(getApplicationContext(), "Please enter mobile number and message", Toast.LENGTH_SHORT).show();
        }
    }


    public void ReadExcel() throws IOException {
        file = new File(FilePath);   //creating a new file instance

        FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);  //creating a Sheet object to retrieve object
//		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
//		style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();


        //		****************Column Value Finder**********************************
        first:
        try {
//			***********first**************

            //for(int i=1;i<=rowCount;i++) {
            for (int i = 1; i <= 1; i++) {
                int cellcount = sheet.getRow(i).getLastCellNum();

                //iterate over each cell to print its value
                // System.out.println("Row"+ i +" data is :");

                for (int j = 0; j <= cellcount; j++) {

                    Row row = sheet.getRow(i);
                    Cell cell = row.getCell(j);
//			    	CellType type = cell.getCellType();

//			************second**************


                    if (cell == null || cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        second:
                        break second;
                    } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        String a = cell.getStringCellValue();


                        if (a.equalsIgnoreCase("Final")) {
                            cellnumber = cell.getColumnIndex();
                            System.out.println("Final" + cellnumber);
                            continue;

                        }
                    } else {
                        break;
                    }


                }

            }
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

//		******************Row Value Finder********************

        try {
//			***********first**************

            first:
            for (int i = 1; i <= rowCount; i++) {
                int cellcount = sheet.getRow(i).getLastCellNum();

                //iterate over each cell to print its value
                // System.out.println("Row"+ i +" data is :");
                second:
                for (int j = 1; j <= cellcount; j++) {

                    Row row = sheet.getRow(i);
                    Cell cell = row.getCell(j);
//			    	CellType type = cell.getCellType();

//			************second**************


                    if (cell == null) {
                        break second;
                    } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        String a = cell.getStringCellValue();


                        if (a.equalsIgnoreCase("Total")) {
                            rownumber = cell.getRowIndex()-1;
                            System.out.println("Total" + rownumber);
                            break first;

                        } else {
                            break second;
                        }


                    }
                }
            }
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }


//		******************Send Message Functionality***************************************


        //last:

        for (int i = 1; i <= rownumber; i++)
            instance:
                    {
                        int cellcount = sheet.getRow(i).getLastCellNum();
                        //int cellcount=sheet.getRow(i).getLastCellNum();
                        ArrayList obj1 = new ArrayList();
                        ArrayList MobileNumber = new ArrayList();
                        //iterate over each cell to print its value
                        // System.out.println("Row"+ i +" data is :");
                        //   System.out.print(cellnumber);


                        second:
                        for (int j = 0; j <= cellnumber; j++) {
                            //  System.out.print(cellnumber);

                            Row row = sheet.getRow(i);
                            Cell cell = row.getCell(j);
//           	CellType type = cell.getCellType();


                            last:
                            if (j == 0 && cell == null || row == null) {
                                break instance;
                            } else if (j == 0 && i != rownumber && cell.getCellType() != Cell.CELL_TYPE_NUMERIC) {
                               break instance;
                            } else if (j == 0 && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

                                DecimalFormat df = new DecimalFormat("#" + "");
                                String num = df.format(cell.getNumericCellValue());
                                MobileNumber.add(num);


                            } else if (j > 0 && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_STRING:
                                        //System.out.print(sheet.getRow(i).getCell(j) +" ");
                                        obj1.add(sheet.getRow(i).getCell(j) + "");
                                }
                            } else if (j > 0 && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_NUMERIC:
                                        DecimalFormat df = new DecimalFormat("#.##" + "");
                                        //df.setMaximumFractionDigits(5);

                                        //System.out.print(df.format(cell.getNumericCellValue()));
                                        //System.out.print(",");
                                        String num = df.format(cell.getNumericCellValue());

                                        //	obj1.add(sheet.getRow(0).getCell(j)+":");
                                        obj1.add(sheet.getRow(1).getCell(j) + ": " + num + "");

                                }
                            } else if (j > 0 && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

                                switch (evaluator.evaluateFormulaCell(cell)) {
                                    case Cell.CELL_TYPE_BOOLEAN:
//        		            System.out.print(cell.getBooleanCellValue());
//        		            System.out.print(",");
                                        obj1.add(sheet.getRow(1).getCell(j));
                                        break;
                                    case Cell.CELL_TYPE_NUMERIC:
                                        DecimalFormat df = new DecimalFormat("#.##" + "");
                                        //df.setMaximumFractionDigits(5);
                                        // System.out.print(df.format(cell.getNumericCellValue()));
                                        // System.out.print(",");
                                        //obj1.add(df.format(cell.getNumericCellValue()+","));
                                        df.setRoundingMode(RoundingMode.DOWN);
                                        String num = df.format(cell.getNumericCellValue());
                                        //obj1.add();
                                        obj1.add(sheet.getRow(1).getCell(j) + ": " + num + "");
                                        //
                                        break;
                                    case Cell.CELL_TYPE_STRING:
//        		            System.out.print(cell.getStringCellValue());
//        		            System.out.print(",");

                                        obj1.add(sheet.getRow(1).getCell(j));
                                        break;
                                }
                            } else if (j > 0 && cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_BLANK:
                                        if (j < cellnumber) {
                                            System.out.print("");
                                        } else {
                                            continue;
                                        }
                                        break;

                                }
                            } else {
                                break;
                            }
                        }

                        StringBuilder buider = new StringBuilder();
                        for(Object value : MobileNumber){
                            buider.append(value);
                        }
                        String PhoneNumber =buider.toString();
                        //System.out.println(PhoneNumber);
                        //number.getEditText().setText(PhoneNumber);
                       // number.getEditText().setText("8128238868");
                        MobNo.add(PhoneNumber);

                        String list = Arrays.toString(obj1.toArray()).replace("[", "").replace("]", "").replace(",", "");
                        	//System.out.println("Hi, "+list);
                        	//message.getEditText().setText("Hi ,"+list);
                        DelMessage.add(list);


                    }


    }

}
