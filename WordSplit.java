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
                
                /* XWPFDocumentをインスタンス化 */
                /* XWPFDocumentをインスタンス化する際に引数の指定が必要 */
                /* FileInputStream(name:パス名)で開くファイルを指定する */
		XWPFDocument doc = new XWPFDocument (new FileInputStream( "C:\\Users\\user\\Desktop\\学習用フォルダ\\pleiades\\学習用フォルダ\\Test\\Test.docx"));
		
                /* XWPFWordExtractorをインスタンス化する(ファイルのテキストを取得するためのクラス) */
                /* 引数としてXWPFDocumentで指定したドキュメントを引数とする */
                XWPFWordExtractor ex = new XWPFWordExtractor(doc);
                
                /* createParagraphを使用するためのインスタンス化 */
		XWPFDocument exDoc = new XWPFDocument ();
                /* createParagraphはドキュメントに段落を追加するためメソッド */
		XWPFParagraph para = exDoc.createParagraph();
                
                /* テキストを設定する */
		XWPFRun run = para.createRun();

                /* ドキュメントから全てのテキストを取得する */
		String tx = ex.getText();

                /* 改行毎で文字列を区切り、文字列をtxsに取得する */
		String [] txs = tx.split("\n");
                
                /* StringBuilder：文字列操作を行うためのクラス */
		StringBuilder sb = new StringBuilder();
                
                /* 文字列の最後尾を上限としてfor文を回す Length-1が上限 */
		for ( int i = 0; i < txs.length; i++) {
			sb = new StringBuilder();
                        
                        /* 行番号+：+行毎の文字列を取得する */
			sb.append((i + 1) + " : " + txs[i]);
                        
                        /* appendで取得したテキストを保持する */
			run.setText(new String(sb));
                        /* 改行を入れる */
			run.addCarriageReturn();
		}

                /* WORDファイルへの書き込み。テキストビューに得た物をFileOutputStreamで指定したwordファイルに書き込む */
		exDoc.write(new FileOutputStream("C:\\Users\\user\\Desktop\\学習用フォルダ\\pleiades\\学習用フォルダ\\Test\\Test_After.docx"));
	}
}