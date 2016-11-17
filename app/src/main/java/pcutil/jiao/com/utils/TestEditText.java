package pcutil.jiao.com.utils;

import android.app.Activity;
import android.os.Bundle;
import android.text.Editable;
import android.text.TextWatcher;
import android.widget.EditText;

/**
 * Created by fxjiao on 16/11/9.
 */

public class TestEditText extends Activity {
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_test);
        testTwoPointEdittext();
    }

    EditText editMimiDiscount;
    private boolean bAdjustSeletion = false;
    public void testTwoPointEdittext(){
        editMimiDiscount = (EditText)findViewById(R.id.edit_test);
        editMimiDiscount.addTextChangedListener(new TextWatcher() {
            @Override
            public void beforeTextChanged(CharSequence s, int start, int count, int after) {
            }

            @Override
            public void onTextChanged(CharSequence s, int start, int before, int count) {
                if (s != null && !s.toString().trim().equals("")) {
                    String input    = s.toString();
                    String[] inputs = input.split("\\.");
                    if(inputs.length > 1){
                        if(inputs[1].length() > 2){
                            bAdjustSeletion = true;
                            editMimiDiscount.setText(inputs[0]+"."+inputs[1].substring(0,2));
                            return;
                        }
                    }
                }
                if(bAdjustSeletion){
                    bAdjustSeletion = false;
                    editMimiDiscount.setSelection(s.length());
                }
            }

            @Override
            public void afterTextChanged(Editable s) {
            }
        });
    }
}
