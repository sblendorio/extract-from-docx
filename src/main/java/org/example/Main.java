package org.example;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.*;

import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class Main {
    public static void main(String[] args) throws Exception {
        String TAG_START = "CONDIZIONI SPECIALI DI ASSICURAZIONE";
        String TAG_END = "Data";
        InputStream is = Main.class.getClassLoader().getResourceAsStream("document.docx");
        ZipSecureFile.setMinInflateRatio(-1.0d);
        XWPFDocument document = new XWPFDocument(is);

        /*
        var numbering = document.getNumbering();
        for (XWPFNum item: numbering.getNums()) {
            XWPFNumbering n = item.getNumbering();
            System.out.println(n);
        }
        System.exit(0);
         */

        List<XWPFParagraph> paraOrigin = document.getParagraphs();
        List<XWPFParagraph> paraFiltered = new LinkedList<>();
        boolean startCopy = false;

        for (XWPFParagraph para : paraOrigin) {
            if (para.getText().contains(TAG_START)) {
                startCopy = true;
                continue;
            }
            if (startCopy && para.getText().contains(TAG_END)) {
                break;
            }
            if (startCopy && para.getParagraphText().trim()!="") {
                paraFiltered.add(para);
            }
        }
        paraFiltered.forEach(x -> {
            System.out.print("{");
            System.out.printf("%5d / ",
                    x.getIndentationHanging()
            );
            System.out.printf("%5d / ",
                    x.getIndentationLeft()
            );
            System.out.printf("%5d / ",
                    x.getNumID()
            );
            System.out.printf("%-5b / ",
                    isTab(x)
            );
            int lvl = -10;
            try {
                lvl = x.getCTP().getPPr().getNumPr().getIlvl().getVal().intValue();
            } catch (Exception e) {
                // Nothing
            }
            System.out.printf("%5d",
                    lvl
            );
            System.out.print("} ");
            System.out.println(x.getParagraphText());
        });

        /*
        -1,-1,>=0 = livello 1
        x,x+k, null = livello 1
        -1, >0, >=0 = livello 2
        x,x,_,false = livello 2
        x,x,_,true = livello 2
        x,x+k,notnull = livello 3


       -1,>0, null = livello 4
         */

        System.out.println("---------------------------------------------------------------");
        paraFiltered.forEach(x -> {
            var v1 = x.getIndentationHanging();
            var v2 = x.getIndentationLeft();
            var v3 = x.getNumID();

            var level = -1;
            if (v1<0 && v2<0 && v3 != null && v3.intValue()>=0)
                level = 1;
            else if (v1>0 && v2>v1 && v3 == null)
                level = 1;
            else if (v1<0 && v2>0 && v3 != null && v3.intValue() >= 0)
                level = 2;
            else if (v1==v2 && !isTab(x))
                level = 2;
            else if (v1==v2 && isTab(x))
                level = 3;
            else if (v1>0 && v2>v1 && v3 != null)
                level = 3;
            else if (v1<0 && v2>0 && v3 == null)
                level = 4;
            else
                level = 5;

            System.out.print("{"+level+"} ");
            System.out.println(x.getParagraphText());
        });

    }

    public static boolean isTab(XWPFParagraph para) {
        return para != null && para.getParagraphText() != null && para.getParagraphText().startsWith("\t");
    }
}
