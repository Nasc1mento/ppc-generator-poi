package br.edu.ifpe;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public abstract class Utils {
    public static void saveDocxFile(XWPFDocument doc, String name) {
        final String path = "build/generated/docs/";
        File f = new File(path);
        if (!f.exists())
            f.mkdirs();

        try (FileOutputStream fos = new FileOutputStream(path+name+".docx")) {
            doc.write(fos);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
