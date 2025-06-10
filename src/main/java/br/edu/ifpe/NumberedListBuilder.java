package br.edu.ifpe;

import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;

public class NumberedListBuilder {

    private final XWPFDocument doc;
    private final BigInteger numId;
    private double baseIndentLevel;

    public NumberedListBuilder(XWPFDocument doc, int levels, double baseIndentLevel) {
        this.doc = doc;
        this.baseIndentLevel = baseIndentLevel;
        this.numId = createNumbering(levels);
    }

    public XWPFParagraph addItem(String text, int level) {
        XWPFParagraph paragraph = doc.createParagraph();
        paragraph.createRun().setText(text);
        paragraph.setNumID(numId);
        CTDecimalNumber ilvl = paragraph.getCTP().getPPr().getNumPr().addNewIlvl();
        ilvl.setVal(BigInteger.valueOf(level));
        return paragraph;
    }

    private BigInteger createNumbering(int levels) {
        CTAbstractNum abstractNum = CTAbstractNum.Factory.newInstance();
        abstractNum.setAbstractNumId(BigInteger.valueOf(0));

        for (int i = 0; i < levels; i++) {
            CTLvl ctLvl = abstractNum.addNewLvl();
            ctLvl.setIlvl(BigInteger.valueOf(i));
            if (i > 0) {
                CTInd ctInd = ctLvl.addNewPPr().addNewInd();
                ctInd.setLeft(Utils.inchesToTwips(this.baseIndentLevel * i));
            }

            ctLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);

            StringBuilder pattern = new StringBuilder();
            for (int j = 0; j <= i; j++) {
                pattern.append("%").append(j + 1).append(".");
            }
            ctLvl.addNewLvlText().setVal(pattern.toString());
            ctLvl.addNewStart().setVal(BigInteger.ONE);
        }

        XWPFNumbering numbering = doc.createNumbering();
        BigInteger abstractNumID = numbering.addAbstractNum(new XWPFAbstractNum(abstractNum));
        return numbering.addNum(abstractNumID);
    }

    public BigInteger getNumId() {
        return this.numId;
    }
}
