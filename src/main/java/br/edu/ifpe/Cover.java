package br.edu.ifpe;

import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

public class Cover implements SectionDoc {

    @Override
    public void write(XWPFDocument doc) {

        String brasaoPath = "brasao_colorido.png";

        XWPFStyles styles = doc.createStyles();
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

        XWPFParagraph headerP = doc.createParagraph();
        headerP.setStyle("defaultStyle");

        RunBuilder rb = new RunBuilder(headerP)
                .img(brasaoPath, PictureType.findByOoxmlId(XWPFDocument.PICTURE_TYPE_PNG), 120, 75)
                .textAndBreak(null)
                .newRun()
                .textAndBreak("MINISTÉRIO DA EDUCAÇÃO")
                .textAndBreak("SECRETARIA DE EDUCAÇÃO PROFISSIONAL E TECNOLÓGICA")
                .textAndBreak("INSTITUTO FEDERAL DE EDUCAÇÃO, CIÊNCIA E TECNOLOGIA DE PERNAMBUCO")
                .newRun()
                .text("CAMPUS ")
                .italic(true)
                .newRun()
                .textAndBreak("IGARASSU")
                .text("DIREÇÃO DE ENSINO");

        for (int i = 0; i < 18; i++) {
            rb.textAndBreak(null);
        }

        rb.newRun()
                .text("PROJETO PEDAGÓGICO DO CURSO SUPERIOR DE TECNOLOGIA EM SISTEMAS PARA INTERNET")
                .fontSize(12)
                .bold(true);


        XWPFHeaderFooterPolicy policy = doc.getHeaderFooterPolicy();
        if (policy == null) {
            policy = doc.createHeaderFooterPolicy();
        }

        XWPFFooter footer = policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph footerP = footer.createParagraph();
        footerP.setStyle("defaultStyle");
        new RunBuilder(footerP)
                .textAndBreak("IGARASSU")
                .text("2025");
    }

    public static void main(String[] args) {
        XWPFDocument doc = new XWPFDocument();
        Cover c = new Cover();
        c.write(doc);

        try (FileOutputStream fos = new FileOutputStream("build/reports/capa.docx")) {
            doc.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
