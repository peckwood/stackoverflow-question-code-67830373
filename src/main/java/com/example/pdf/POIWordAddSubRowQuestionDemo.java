package com.example.pdf;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * 替换表格中的文字
 */
public class POIWordAddSubRowQuestionDemo{
    public static void main(String[] args) throws IOException, XmlException{
        ClassLoader classLoader = POIWordAddSubRowQuestionDemo.class.getClassLoader();
        InputStream inputStream = classLoader.getResourceAsStream("input.docx");
        String outputDocxPath = "F:/TEMP/output.docx";
        assert inputStream != null;
        XWPFDocument doc = new XWPFDocument(inputStream);
        XWPFTable table = doc.getTables().get(0);
        //this is 'old row 2'
        XWPFTableRow secondRow = table.getRows().get(1);

        //create a new row that is based on 'old row 2'
        CTRow ctrow = CTRow.Factory.parse(secondRow.getCtRow().newInputStream());
        XWPFTableRow newRow = new XWPFTableRow(ctrow, table);
        XWPFRun xwpfRun = newRow.getCell(1).getParagraphs().get(0).getRuns().get(0);
        //set row text
        xwpfRun.setText("new row", 0);
        // add new row below 'old row 2'
        table.addRow(newRow, 2);

        //merge cells at first column of 'old row 2', 'new row', and 'old row 3'
        mergeCellVertically(doc.getTables().get(0), 0, 1, 3);

        FileOutputStream fos = new FileOutputStream(outputDocxPath);
        doc.write(fos);
        fos.close();
    }

    static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for(int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTVMerge vmerge = CTVMerge.Factory.newInstance();
            if(rowIndex == fromRow){
                // The first merged cell is set with RESTART merge value
                vmerge.setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                vmerge.setVal(STMerge.CONTINUE);
                // and the content should be removed
                for (int i = cell.getParagraphs().size(); i > 0; i--) {
                    cell.removeParagraph(0);
                }
                cell.addParagraph();
            }
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
            tcPr.setVMerge(vmerge);
        }
    }
}
