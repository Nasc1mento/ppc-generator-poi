package br.edu.ifpe;

import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;


public class TableOfContents implements SectionDoc {

    private static final double BASE_INDENT_VALUE = 0.5;

    @Override
    public void write(XWPFDocument doc) {
        BigInteger numId = getNumId(doc);

        createNumberedParagraph(doc, numId, "Subject 01", BigInteger.ZERO);
        createNumberedParagraph(doc, numId, "Item 01", BigInteger.ONE);
        createNumberedParagraph(doc, numId, "Item 02", BigInteger.ONE);
        createNumberedParagraph(doc, numId, "Subject 02", BigInteger.ZERO);
        createNumberedParagraph(doc, numId, "Item 03", BigInteger.ONE);
        createNumberedParagraph(doc, numId, "Item 04", BigInteger.ONE);
        createNumberedParagraph(doc, numId, "Item 05", BigInteger.TWO);
        createNumberedParagraph(doc, numId, "Item 06", BigInteger.TWO);
    }

    public static void main(String[] args){

        XWPFDocument doc = new XWPFDocument();
        TableOfContents toc = new TableOfContents();
        toc.write(doc);

        try (FileOutputStream fos = new FileOutputStream("build/reports/toc.docx")) {
            doc.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static BigInteger getNumId(XWPFDocument document) {
        CTAbstractNum abstractNum = CTAbstractNum.Factory.newInstance();
        abstractNum.setAbstractNumId(BigInteger.valueOf(0));

        for (int i = 1; i <= 3; i++) {
            createLevel(abstractNum, i);
        }

        XWPFNumbering numbering = document.createNumbering();
        BigInteger abstractNumID = numbering.addAbstractNum(new XWPFAbstractNum(abstractNum));
        return numbering.addNum(abstractNumID);
    }

    public static void createLevel(CTAbstractNum ctAbstractNum, int lvl) {
        CTLvl ctLvl = ctAbstractNum.addNewLvl();
        ctLvl.setIlvl(BigInteger.valueOf(lvl - 1));
        if (lvl != 1) {
            CTInd ctInd = ctLvl.addNewPPr().addNewInd();
            ctInd.setLeft(inchesToTwips((BASE_INDENT_VALUE * (lvl - 1))));
        }

        ctLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
        StringBuilder sb = new StringBuilder();

        for (int i = 1; i <= lvl; i++) {
            sb.append("%").append(i).append(".");
        }

        ctLvl.addNewLvlText().setVal(sb.toString());
        ctLvl.addNewStart().setVal(BigInteger.ONE);
    }

    private static BigInteger inchesToTwips(double inches) {
        return BigInteger.valueOf((long) (1440L * inches));
    }

    private static void createNumberedParagraph(XWPFDocument doc, BigInteger numId, String paragraphText, BigInteger numLevel) {
        XWPFParagraph paragraph = doc.createParagraph();
        paragraph.createRun().setText(paragraphText);
        paragraph.setNumID(numId);
        CTDecimalNumber ctDecimalNumber = paragraph.getCTP().getPPr().getNumPr().addNewIlvl();
        ctDecimalNumber.setVal(numLevel);
    }
}