package com.example.doc;

import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Color;
import android.graphics.Typeface;
import android.os.Bundle;
import android.text.SpannableString;
import android.text.method.ScrollingMovementMethod;
import android.text.style.UnderlineSpan;
import android.util.DisplayMetrics;
import android.util.Log;
import android.view.Gravity;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.view.View;
import android.widget.ImageView;
import android.widget.LinearLayout;
import android.widget.RelativeLayout;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {


    LinearLayout mainUI;

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

        mainUI = findViewById(R.id.MainUI);

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

//                        WordExtractor extractor = new WordExtractor(wordDoc);
//
//                        for(String rawText : extractor.getParagraphText()) {
//                            String text = extractor.stripFields(rawText);
//                            Log.i("Text",text);
//                        }

                        WordExtractor extractor = new WordExtractor(wordDoc);

                        Range range = wordDoc.getRange();
                        String[] paragraphs = extractor.getParagraphText();

                        PicturesTable picturesTable = wordDoc.getPicturesTable();
                        List<Picture> all = picturesTable.getAllPictures();

                        for(int i =0;i < paragraphs.length;i++){
                            Paragraph pr = range.getParagraph(i);

//                            Log.i("text",pr.text());
                            int j =0 ;

                            while(true){
                                CharacterRun run = pr.getCharacterRun(j++);

                                StyleDescription style = wordDoc.getStyleSheet().getStyleDescription(run.getSubSuperScriptIndex());
                                String styleName = style.getName();
                                String font = run.getFontName();
                                int size = run.getFontSize();
                                String paraText = pr.text();
                                Boolean b = run.isBold();
                                int u = run.getUnderlineCode();

                                addTextViews(paraText,size,b,u,font);

                                if(picturesTable.hasPicture(run)){
                                    Picture p = picturesTable.extractPicture(run,true);
                                    traversePictures(p);
                                }

                                Log.i("name",styleName);
                                Log.i("font",Integer.toString(size));
                                Log.i("family",font);
                                Log.i("text",paraText);


//
//                                List<Picture> pictures = wordDoc.getPicturesTable().getAllPictures();
//                                traversePictures(pictures);

                                if (run.getEndOffset() == pr.getEndOffset()) {
                                    break;
                                }
                            }
                        }

//                        extractImages(wordDoc);


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

    private void traversePictures(Picture pic) {

        Log.i("pictureData",pic.getContent().toString());

        DisplayMetrics displayMetrics = new DisplayMetrics();
        getWindowManager().getDefaultDisplay().getMetrics(displayMetrics);
        int height = displayMetrics.heightPixels;
        int width = displayMetrics.widthPixels;


        long w = pic.getWidth();
        long h = pic.getHeight();

        addElements(pic,w,h,height,width);
    }

    int TagCnt = 0;

    public void addTextViews(String content, int s, Boolean b,int u,String f) {
        TextView text = new TextView(this);
        text.setLayoutParams(new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT,RelativeLayout.LayoutParams.WRAP_CONTENT));
        SpannableString c = new SpannableString(content);
        c.setSpan(new UnderlineSpan(),0,content.length(),0);
        text.setTag(TagCnt);

        ArrayList<TextView> textViews = new ArrayList<>();
//        TextView myText = (TextView) mainUI.findViewWithTag(TagCnt - i); // get the element

        for (int i=TagCnt - 1;i >= 0;i--){

            TextView myText = (TextView) mainUI.findViewWithTag(i); // get the element
            if(myText != null) {
                Log.d("CONT ",String.valueOf(content));
                if((content.equals(myText.getText().toString()))){
                    Log.d("MYTXT ",myText.getText().toString());
                    Log.d("ICNT ",String.valueOf(i));
                    break;

                }else{
                    setProperText(text, content, s, b, u, f);
                    break;
                }
            }else{
                Log.d("NTAGCNT ",String.valueOf(TagCnt));
                Log.d("NICNT ",String.valueOf(i));
                Log.d("NULLTXT ",content);
            }
        }
        if(TagCnt == 0 ){
            setProperText(text, content, s, b, u, f);
        }
        TagCnt = TagCnt + 1;
    }

    public void setProperText(TextView text,String content, int s, Boolean b, int u,String fontFamily) {
        SpannableString c = new SpannableString(content);
        c.setSpan(new UnderlineSpan(), 0, content.length(), 0);
        text.setPadding(50, 10, 50, 10);
        text.setTextAlignment(View.TEXT_ALIGNMENT_VIEW_START);

        if (b) {
            text.setTextColor(Color.BLACK);
            text.setTextSize((int) ((3 * s) / 4));
            //text.setTypeface(null, FontStyle.fontFamily);
            text.setTypeface(null, Typeface.BOLD);

            if (u > 0) {
                text.setText(c);
            } else {
                //text.setTextAlignment(View.TEXT_ALIGNMENT_CENTER);
                text.setText(content);
//                text.setGravity(Gravity.CENTER);
            }
//                if(!c.equals(myText.getText().toString()) || !content.equals(myText.getText().toString())) {
            mainUI.addView(text);
//                }
        } else {
            text.setTextColor(Color.BLACK);
            text.setTextSize((int) ((3 * s) / 4));

            if (u > 0) {
                text.setText(c);
            } else {
                //               text.setTextAlignment(View.TEXT_ALIGNMENT_CENTER);
                text.setText(content);
//                text.setGravity(Gravity.CENTER);
            }
//                if(!c.equals(myText.getText().toString()) || !content.equals(myText.getText().toString())) {
            mainUI.addView(text);
//                }
        }

        // Adds the view to the layout
        LinearLayout textLayout = new LinearLayout(getApplicationContext());
        RelativeLayout.LayoutParams params = new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT, RelativeLayout.LayoutParams.WRAP_CONTENT);
        params.addRule(RelativeLayout.BELOW, text.getId());
        textLayout.setLayoutParams(params);
        mainUI.addView(textLayout);

    }

        public  void addElements(Picture pictureData,long w,long h,int height,int width){

        ArrayList<ImageView> imageViews = new ArrayList<>();
//
        ImageView image = new ImageView(this);

        if((int)w >= width){
            w = width - 90;
            h = h / (width-100);
        }
        else if((int)w < width){
            w = width ;
            h = height / 3;
//            h = height;
        }

//        image.setForegroundGravity(Gravity.CENTER);
        image.setPadding(50,10,50,10);
        image.setLayoutParams(new RelativeLayout.LayoutParams((int)w,(int)h));
        image.setMaxHeight((int)h);
        image.setMaxWidth((int)w);
//        image.setAdjustViewBounds(true);
//        image.setScaleType(ImageView.ScaleType.MATRIX);
        InputStream inputStream = new ByteArrayInputStream(pictureData.getContent());
        Bitmap bitmap = BitmapFactory.decodeStream(inputStream);

//        Matrix matrix = new Matrix();
//        matrix.postScale(1/2800, 1/2800);
//        Bitmap scaledBitmap = Bitmap.createBitmap(bitmap,0,0,(int)w,(int)h,matrix,true);

        image.setImageBitmap(bitmap);
        mainUI.addView(image);


        // Adds the view to the layout
        LinearLayout imageLayout = new LinearLayout(getApplicationContext());
        RelativeLayout.LayoutParams params = new RelativeLayout.LayoutParams(RelativeLayout.LayoutParams.WRAP_CONTENT, RelativeLayout.LayoutParams.WRAP_CONTENT);
        params.addRule(RelativeLayout.BELOW, image.getId());
        imageLayout.setLayoutParams(params);
        mainUI.addView(imageLayout);

    }

}
