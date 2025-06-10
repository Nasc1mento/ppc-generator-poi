package br.edu.ifpe;

import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.IOException;
import java.io.InputStream;
import java.util.Objects;


public class RunBuilder {
    private final XWPFParagraph paragraph;
    private XWPFRun run;

    public RunBuilder(XWPFParagraph paragraph) {
        Objects.requireNonNull(paragraph);
        this.paragraph = paragraph;
        this.run = this.newRun().build();
    }

    public RunBuilder newRun() {
        this.run = this.paragraph.createRun();
        return this;
    }

    public RunBuilder text(String text) {
        this.run.setText(text);
        return this;
    }

    public RunBuilder textln(String text) {
        this.run.setText(text);
        this.run.addBreak();
        return this;
    }

    public RunBuilder ln() {
        this.run.addBreak();
        return this;
    }

    public RunBuilder italic(boolean value) {
        this.run.setItalic(value);
        return this;
    }

    public RunBuilder bold(boolean value) {
        this.run.setBold(value);
        return this;
    }

    public RunBuilder underline(UnderlinePatterns p) {
        this.run.setUnderline(p);
        return this;
    }

    public RunBuilder fontSize(int size) {
        this.run.setFontSize(size);
        return this;
    }

    public RunBuilder fontFamily(String font) {
        this.run.setFontFamily(font);
        return this;
    }

    public RunBuilder color(String hexStr) {
        this.run.setColor(hexStr);
        return this;
    }

    public XWPFRun build() {
        return this.run;
    }

    public RunBuilder img(String path, PictureType type, int width, int height)  {
        Objects.requireNonNull(path);
        Objects.requireNonNull(type);
        try (InputStream is = getClass().getClassLoader().getResourceAsStream(path)) {
            this.run.addPicture(is, type, path, Units.toEMU(width), Units.toEMU(height));
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }

        return this;
    }
}
