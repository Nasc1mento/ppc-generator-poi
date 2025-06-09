package br.edu.ifpe;

import org.apache.poi.xwpf.usermodel.XWPFRun;

public abstract class Utils {


    public static XWPFRun setTextAndBreak(String text, XWPFRun run) {
        run.setText(text);
        run.addBreak();
        return run;
    }
}
