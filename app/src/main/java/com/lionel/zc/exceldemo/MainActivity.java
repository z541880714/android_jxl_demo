package com.lionel.zc.exceldemo;

import androidx.appcompat.app.AppCompatActivity;

import android.os.Bundle;
import android.util.Log;
import android.view.View;

public class MainActivity extends AppCompatActivity {

    ExcelManager em = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        em = new ExcelManager(this.getApplicationContext());
    }


    public void onClick(View view) {
        Log.i("log_zc", "onClick: aaaa");
        if (em == null) {
            return;
        }
        em.writeExcel();
    }
}