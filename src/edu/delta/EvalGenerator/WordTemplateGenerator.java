/**
 * This class is responsible for generating the template comments
 * sheets that the comments are written in. 
 */

package edu.delta.EvalGenerator;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class WordTemplateGenerator {
	public WordTemplateGenerator(){
		
	}
	
	/**
	 * Generates Word template sheet based on a DocInfo object. 
	 */
	public void generateWordTemplate(DocInfo info, File saveLoc){
		System.out.println("CREATING WORD DOC");
		System.out.println("GenerateWordDoc path: " + saveLoc.getPath());
		WordprocessingMLPackage wordMLPackage;
		try {
			wordMLPackage = WordprocessingMLPackage.createPackage();
			wordMLPackage.getMainDocumentPart().addParagraphOfText("Josh");
			wordMLPackage.getMainDocumentPart().addParagraphOfText("Klingler");
			wordMLPackage.save(new java.io.File(saveLoc.getPath() + "\\HelloWord1.docx"));
			System.out.println("Saved to " + saveLoc.getPath() + "\\HelloWord1.docx");
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		} 
		
//		try {
//			 
//			String content = "This is the content to write into file";
// 
//			System.out.println(saveLoc.getAbsolutePath() + "\\fileName.txt");
//			File file = new File("C:/Users/joshuaklingler/desktop/test.txt");
// 
//			// if file doesn't exists, then create it
//			if (!file.exists()) {
//				file.createNewFile();
//			}
// 
//			FileWriter fw = new FileWriter(file.getAbsoluteFile());
//			BufferedWriter bw = new BufferedWriter(fw);
//			bw.write(content);
//			bw.close();
// 
//			System.out.println("Done");
// 
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
	}
	
}
