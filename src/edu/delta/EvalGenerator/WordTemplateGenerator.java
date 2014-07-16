/**
 * This class is responsible for generating the template comments
 * sheets that the comments are written in. 
 */

package edu.delta.EvalGenerator;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class WordTemplateGenerator {
	public WordTemplateGenerator(){
		System.out.println("CREATING WORD DOC");
		WordprocessingMLPackage wordMLPackage;
		try {
			wordMLPackage = WordprocessingMLPackage.createPackage();
			wordMLPackage.getMainDocumentPart().addParagraphOfText("Josh");
			wordMLPackage.save(new java.io.File("src/HelloWord1.docx"));
			System.out.println("done");
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
}
