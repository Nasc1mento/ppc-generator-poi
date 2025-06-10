package br.edu.ifpe;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;


public class TableOfContents implements DocumentStep {

    private final XWPFDocument doc;

    public TableOfContents(XWPFDocument doc) {
        this.doc = doc;
    }

    @Override
    public void run() {

        XWPFStyles styles = this.doc.createStyles();
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId("defaultStyle");
        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal("defaultStyle");
        ctStyle.setName(styleName);
        CTRPr rPr = ctStyle.addNewRPr();
        rPr.addNewSz().setVal(BigInteger.valueOf(20));
        CTFonts ctFonts = rPr.addNewRFonts();
        ctFonts.setAscii("Arial");
        ctFonts.setHAnsi("Arial");
        ctFonts.setCs("Arial");
        CTPPrGeneral pPr = ctStyle.addNewPPr();
        pPr.addNewJc().setVal(STJc.CENTER);
        XWPFStyle dfStyle = new XWPFStyle(ctStyle);
        dfStyle.setType(STStyleType.PARAGRAPH);
        styles.addStyle(dfStyle);

        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);

        new RunBuilder(title)
                .textln("SUMÁRIO")
                .fontFamily("Arial")
                .bold(true);


        NumberedListBuilder listBuilder = new NumberedListBuilder(doc, 3, 0.5);

        listBuilder.addItem("DADOS DE IDENTIFICAÇÃO", 0);
        listBuilder.addItem("Da Mantenedora", 1);
        listBuilder.addItem("Da Instituição", 1);
        listBuilder.addItem("ORGANIZAÇÃO DIDÁTICO-PEDAGÓGICA ", 0);
        listBuilder.addItem("Histórico da Instituição", 1);
        listBuilder.addItem("O IFPE Campus Igarassu", 2);

        
    }

    public static void main(String[] args) {

        XWPFDocument doc = new XWPFDocument();
        TableOfContents toc = new TableOfContents(doc);
        toc.run();

        Utils.saveDocxFile(doc, "toc");

    }
}