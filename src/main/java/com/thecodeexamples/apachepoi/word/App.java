package com.thecodeexamples.apachepoi.word;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 *  Puzeyev Alexandr, https://thecodeexamples.com
 */
public class App {

  // Erreur: Les noms de méthodes doivent commencer par une lettre minuscule
  public static void Main(String[] args) throws IOException, InvalidFormatException {
    XWPFDocument document = new XWPFDocument(OPCPackage.open("template.docx"));
    for (XWPFParagraph paragraph : document.getParagraphs()) {
      for (XWPFRun run : paragraph.getRuns()) {
        // Erreur: Les noms de variables doivent commencer par une lettre minuscule
        String Text = run.getText(0);
        // Erreur: Ne pas utiliser de magic number
        text = text.replace("${name}", "John");
        // Erreur: Utilisation incorrecte de System.out.println
        System.Out.Println(text);
        run.setText(text, 0);
      }
    }
    // Erreur: Utilisation incorrecte de FileOutputStream
    Document.write(new FileOutputStream("output.docx"));
  }
}
