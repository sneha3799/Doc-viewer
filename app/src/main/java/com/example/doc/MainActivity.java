package com.example.doc;

import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.os.Bundle;
import android.text.method.ScrollingMovementMethod;
import android.util.Log;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    TextView textView;

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        MenuInflater inflater = getMenuInflater();
        inflater.inflate(R.menu.menu, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        switch (item.getItemId()) {
            case R.id.item:
                openDocumentFromFileManager();
                return true;
//            case R.id.item1:
//                return true;
            default:
                return super.onOptionsItemSelected(item);
        }
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        textView = (TextView) findViewById(R.id.textView);

    }

    private void openDocumentFromFileManager() {
        //this is the action to open doc file from file manager
        Intent i = new Intent();
        i.setType("application/*");
        i.setAction(Intent.ACTION_GET_CONTENT);
        if (PermissionsHelper.getPermission(this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE, R.string.title_storage_permission
                , R.string.text_storage_permission, 1111)) {
            startActivityForResult(Intent.createChooser(i, "Select Document"), 111);
        }
    }


    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        try {
            if (resultCode == RESULT_OK) {
                switch (requestCode) {
                    case 111:
                        //this is action performed after openDocumentFromFileManager() when doc is selected
                        FileInputStream inputStream = (FileInputStream) getContentResolver().openInputStream(data.getData());

                        HWPFDocument wordDoc = new HWPFDocument(inputStream);

                        extractImages(wordDoc);


//                        extractText(docx);

//                        List<XWPFParagraph> paragraphList = docx.getParagraphs();
//
//                        for (XWPFParagraph paragraph:paragraphList){
//                            String paragrapthText = paragraph.getText();
//                            System.out.println(paragrapthText);
//                        }


                }
            } else {
                Toast.makeText(this, "Your file is not loaded", Toast.LENGTH_SHORT).show();
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void extractImages(HWPFDocument doc) {
        try {

            //function to retrieve all the images from the doc file that have been selected
//
//            HWPFDocument wordDoc = new HWPFDocument(new FileInputStream(fileName));
//
//            XWPFWordExtractor extractor = new XWPFWordExtractor(docx);
//
////            extractText(extractor);
//
//            System.out.println(extractor.getText());


            WordExtractor extractor = new WordExtractor(doc);

            textView.setText(extractor.getText());
            textView.setMovementMethod(new ScrollingMovementMethod());

            int pages = doc.getSummaryInformation().getPageCount();
            Log.i("Pages", Integer.toString(pages));

//            List<XWPFPictureData> picList = docx.getAllPackagePictures();
//
//            int i = 1;
//
//            for (XWPFPictureData pic : picList) {
//
//                System.out.print(pic.getPictureType());
//                System.out.print(pic.getData());
//                System.out.println("Image Number: " + i + " " + pic.getFileName());
//                switch (i){
//                    case 1:
//                        InputstreamGeneralPhotos1 = new ByteArrayInputStream(pic.getData());
//                        if (InputstreamGeneralPhotos1 != null){
//                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos1);
//                            img_generalPhotos1.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                        }
//                        break;
//                    case 2:
//                        InputstreamGeneralPhotos2 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 3:
//                        InputstreamGeneralPhotos3 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 4:
//                        InputstreamGeneralPhotos4 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 5:
//                        InputstreamGeneralPhotos5 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 6:
//                        InputstreamGeneralPhotos6 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 7:
//                        InputstreamBicyclePhotos1 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 8:
//                        InputstreamBicyclePhotos2 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 9:
//                        InputstreamFourhoursPhotos1 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 10:
//                        InputstreamFourhoursPhotos2 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 11:
//                        InputstreamFourhoursPhotos3 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 12:
//                        InputstreamFourhoursPhotos4 = new ByteArrayInputStream(pic.getData());
//                        break;
//                }
//                i++;
//            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
