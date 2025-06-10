package br.edu.ifpe;

import org.apache.poi.xwpf.usermodel.XWPFDocument;


public class TableOfContents implements DocumentStep {

    private static final double BASE_INDENT_VALUE = 0.5;

    private final XWPFDocument doc;

    public TableOfContents(XWPFDocument doc) {
        this.doc = doc;
    }

    @Override
    public void run() {
        NumberedListBuilder listBuilder = new NumberedListBuilder(doc, 3, 0.5);

        listBuilder.addItem("Subject 01", 0);
        listBuilder.addItem("Item 01", 1);
        listBuilder.addItem("Item 02", 1);
        listBuilder.addItem("Subject 02", 0);
        listBuilder.addItem("Item 03", 1);
        listBuilder.addItem("Item 04", 1);
        listBuilder.addItem("Item 05", 2);
        listBuilder.addItem("Item 06", 2);
    }

    public static void main(String[] args) {

        XWPFDocument doc = new XWPFDocument();
        TableOfContents toc = new TableOfContents(doc);
        toc.run();

        Utils.saveDocxFile(doc, "toc");

    }
}