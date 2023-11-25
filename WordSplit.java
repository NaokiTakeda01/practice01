/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package trial.mavenproject2;

import java.io.FileInputStream;     /*  */
import java.io.FileOutputStream;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor; 
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordSplit {

	static StringBuilder sb;

	public static void main  (String[] args) throws Exception {
                
                /* FileInputStream(name:パス名)で開くファイルを指定する */
		XWPFDocument doc = new XWPFDocument (new FileInputStream( "C:\\Users\\user\\Desktop\\学習用フォルダ\\pleiades\\学習用フォルダ\\Test\\Test.docx"));
		
                /* ファイルのテキストを取得する */
                XWPFWordExtractor ex = new XWPFWordExtractor(doc);
		XWPFDocument exDoc = new XWPFDocument ();
		XWPFParagraph para = exDoc.createParagraph();

		XWPFRun run = para.createRun();

		String tx = ex.getText();

		String [] txs = tx.split("\n");
		StringBuilder sb = new StringBuilder();
		for ( int i = 0; i < txs.length; i++) {
			sb = new StringBuilder();
			sb.append((i + 1) + " : " + txs[i]);
			run.setText(new String(sb));
			run.addCarriageReturn();
		}

		exDoc.write(new FileOutputStream("C:\\Users\\user\\Desktop\\学習用フォルダ\\pleiades\\学習用フォルダ\\Test\\Test_After.docx"));
	}
}