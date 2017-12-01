package com.example.admin.exporttoexcel;

import android.content.ContentValues;
import android.content.Intent;
import android.database.Cursor;
import android.database.sqlite.SQLiteDatabase;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Locale;

import Database.DatabaseHelper;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class MainActivity extends AppCompatActivity {

    public static String TAG="TAG";
    SQLiteDatabase db;
    private Button btnexporttoexcel,btnexporttopdf;
    private EditText edtname,edtaddress,edtemail,edtno;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        DatabaseHelper databaseHelper=new DatabaseHelper(this,"kuch bhi",null,1);
        db=databaseHelper.getWritableDatabase();

        init();

        btnexporttoexcel.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                insertdata();
                Cursor cursor=db.rawQuery("select * from personinfo",null);
                exporttoexcel(cursor);
            }
        });
        btnexporttopdf.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                insertdata();
//                Intent intent=new Intent(MainActivity.this,PDFActivity.class);
//                startActivity(intent);
                 createpdf();
            }
        });

    }

    private void createpdf() {
        String FILE = Environment.getExternalStorageDirectory().toString()
                + "/PDF/" + edtname.getText().toString()+".pdf";

        // Add Permission into Manifest.xml
        // <uses-permission
        // android:name="android.permission.WRITE_EXTERNAL_STORAGE"/>

        // Create New Blank Document
        Document document = new Document(PageSize.A4);

        // Create Directory in External Storage
        String root = Environment.getExternalStorageDirectory().toString();
        File myDir = new File(root + "/PDF");
        myDir.mkdirs();

        // Create Pdf Writer for Writting into New Created Document
        try {
            PdfWriter.getInstance(document, new FileOutputStream(FILE));

            // Open Document for Writting into document
            document.open();

            // User Define Method
            addMetaData(document);
            addTitlePage(document);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (DocumentException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        // Close Document after writting all content
        document.close();

        Toast.makeText(this, "PDF File is Created. Location : " + FILE,
                Toast.LENGTH_LONG).show();
    }
    public void addMetaData(Document document)

    {
        document.addTitle("RESUME");
        document.addSubject("Person Info");
        document.addKeywords("Personal,	Education, Skills");
        document.addAuthor("TAG");
        document.addCreator("TAG");
        document.bottomMargin();

    }
    public void addTitlePage(Document document) throws DocumentException {
        // Font Style for Document
        Font catFont = new Font(Font.FontFamily.TIMES_ROMAN, 18, Font.BOLD);
        Font titleFont = new Font(Font.FontFamily.TIMES_ROMAN, 22, Font.BOLD);
        Font smallBold = new Font(Font.FontFamily.TIMES_ROMAN, 12, Font.BOLD);
        Font normal = new Font(Font.FontFamily.TIMES_ROMAN, 12, Font.NORMAL);

        // Start New Paragraph
        Paragraph prHead = new Paragraph();
        // Set Font in this Paragraph
        prHead.setFont(titleFont);
        // Add item into Paragraph
        prHead.add(edtname.getText().toString()+"\n");

        // Create Table into Document with 1 Row
        PdfPTable myTable = new PdfPTable(1);
        // 100.0f mean width of table is same as Document size
        myTable.setWidthPercentage(100.0f);


        // Create New Cell into Table
        PdfPCell myCell = new PdfPCell(new Paragraph(""));
        myCell.setBorder(Rectangle.BOTTOM);

        // Add Cell into Table
        myTable.addCell(myCell);

        prHead.setFont(catFont);
      //  prHead.add("\nName1 Name2\n");
        prHead.setAlignment(Element.ALIGN_CENTER);

        // Add all above details into Document
        document.add(prHead);
        document.add(myTable);

        document.add(myTable);

        // Now Start another New Paragraph
        Paragraph prPersinalInfo = new Paragraph();
        prPersinalInfo.setFont(smallBold);
        prPersinalInfo.add("Address\n");
       // prPersinalInfo.add("Address 2\n");
        prPersinalInfo.add(edtaddress.getText().toString()+"\n");
        //prPersinalInfo.add("Country: USA Zip Code: 000001\n");
        prPersinalInfo.add(edtemail.getText().toString()+"\n");
        prPersinalInfo.add(edtno.getText().toString()+"\n");

        prPersinalInfo.setAlignment(Element.ALIGN_CENTER);

        document.add(prPersinalInfo);
        document.add(myTable);

        document.add(myTable);

        Paragraph prProfile = new Paragraph();
        prProfile.setFont(smallBold);
        prProfile.add("\n \n Profile : \n ");
        prProfile.setFont(normal);
        prProfile
                .add("\nI am "+edtname.getText().toString()+ " I am Android Application Developer at Google.");

        prProfile.setFont(smallBold);
        document.add(prProfile);

        // Create new Page in PDF
        document.newPage();
    }



    private void exporttoexcel(Cursor cursor) {

        String personname=edtname.getText().toString();
        final String fileName = personname+".xls";

        File sdCard = Environment.getExternalStorageDirectory();
        File directory = new File(sdCard.getAbsolutePath() + "/DEMOEXCEL");

        if (!directory.isDirectory()) {
            directory.mkdirs();
        }
        File file = new File(directory, fileName);

        WorkbookSettings wbSettings = new WorkbookSettings();

        wbSettings.setLocale(new Locale("en", "EN"));
        WritableWorkbook workbook;

        try {
            workbook = Workbook.createWorkbook(file, wbSettings);
            WritableSheet sheet = workbook.createSheet("MyShoppingList", 0);

            try {

                sheet.addCell(new Label(0, 0, "NAME"));
                sheet.addCell(new Label(1, 0, "ADDRESS"));
                sheet.addCell(new Label(2, 0, "EMAIL"));
                sheet.addCell(new Label(3, 0, "MOBILENO"));

                if (cursor.moveToFirst()) {
                    do {

                        String name = cursor.getString(cursor.getColumnIndex("NAME"));
                        String adddress = cursor.getString(cursor.getColumnIndex("ADDRESS"));
                        String email = cursor.getString(cursor.getColumnIndex("EMAILID"));
                        String mobileno = cursor.getString(cursor.getColumnIndex("PHONENO"));

                        Log.d(TAG, "Name=" + name);
                        Log.d(TAG, "ADDRESS=" + adddress);
                        Log.d(TAG, "Eamilid==" + email);
                        Log.d(TAG, "phoneno==" + mobileno);
                        //String Emailid = cursor.getString(cursor.getColumnIndex("ADDRESS"));
                        // String phonenno = cursor.getString(cursor.getColumnIndex("ADDRESS"));

                        int i = cursor.getPosition() + 1;

                        sheet.addCell(new Label(0, i, name));
                        sheet.addCell(new Label(1, i, adddress));
                        sheet.addCell(new Label(2, i, email));
                        sheet.addCell(new Label(3, i, mobileno));


                    } while (cursor.moveToNext());
                    cursor.close();
                    Toast.makeText(this, "Sucess check on path"+directory+" "+fileName, Toast.LENGTH_SHORT).show();
                }

            } catch (RowsExceededException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            }
            workbook.write();
            try {
                workbook.close();
            } catch (WriteException e) {

            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void init() {
        btnexporttoexcel=(Button)findViewById(R.id.ExportToExcel);
        btnexporttopdf=(Button)findViewById(R.id.ExportToPdf);
        edtname=(EditText)findViewById(R.id.name);
        edtaddress=(EditText)findViewById(R.id.name);
        edtemail=(EditText)findViewById(R.id.Email);
        edtno=(EditText)findViewById(R.id.no);
    }

    public void insertdata(){

      //  db.delete("personinfo",null,null);
        ContentValues contentValues=new ContentValues();
        contentValues.put("NAME",edtname.getText().toString());
        contentValues.put("ADDRESS",edtaddress.getText().toString());
        contentValues.put("EMAILID",edtemail.getText().toString());
        contentValues.put("PHONENO",edtno.getText().toString());

       long insert= db.insert("personinfo",null,contentValues);

        if (insert>0){
            Log.d(TAG,"inserted successfully");
        }
    }
}
