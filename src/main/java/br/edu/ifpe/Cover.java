package br.edu.ifpe;

import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;
import java.util.Objects;

import static org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PNG;

public class Cover implements DocumentStep {

    private static final String BRASAO_IMAGE = "brasao_colorido.png";
    private static final String DEFAULT_STYLE_ID = "coverDefaultStyle";


    private XWPFDocument doc;


    public Cover(XWPFDocument doc) {
        Objects.requireNonNull(doc);
        this.doc = doc;
    }

    @Override
    public void run() {
        XWPFStyles styles = this.doc.createStyles();
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(DEFAULT_STYLE_ID);
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

        XWPFParagraph headerP = this.doc.createParagraph();
        headerP.setStyle(DEFAULT_STYLE_ID);

        RunBuilder rb = new RunBuilder(headerP)
                .img(BRASAO_IMAGE, PictureType.findByOoxmlId(PICTURE_TYPE_PNG), 120, 75).ln()
                .newRun()
                .textln("MINISTÉRIO DA EDUCAÇÃO")
                .textln("SECRETARIA DE EDUCAÇÃO PROFISSIONAL E TECNOLÓGICA")
                .textln("INSTITUTO FEDERAL DE EDUCAÇÃO, CIÊNCIA E TECNOLOGIA DE PERNAMBUCO")
                .newRun()
                .text("CAMPUS ")
                .italic(true)
                .newRun()
                .textln("IGARASSU")
                .text("DIREÇÃO DE ENSINO");

        for (int i = 0; i < 18; i++) {
            rb.ln();
        }

        rb.newRun()
                .text("PROJETO PEDAGÓGICO DO CURSO SUPERIOR DE TECNOLOGIA EM SISTEMAS PARA INTERNET")
                .fontSize(12)
                .bold(true);


        XWPFHeaderFooterPolicy policy = this.doc.getHeaderFooterPolicy();
        if (policy == null) {
            policy = this.doc.createHeaderFooterPolicy();
        }
        XWPFFooter footer = policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph footerP = footer.createParagraph();
        footerP.setStyle(DEFAULT_STYLE_ID);
        new RunBuilder(footerP)
                .textln("IGARASSU")
                .text("2025");
    }

    public static void main(String[] args) {
        XWPFDocument doc = new XWPFDocument();
        Cover c = new Cover(doc);
        c.run();
        Utils.saveDocxFile(doc, "cover");
    }
}
