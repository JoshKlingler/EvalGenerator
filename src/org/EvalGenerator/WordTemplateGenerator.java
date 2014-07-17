/**
 * This class is responsible for generating the template comments
 * sheets that the comments are written in. 
 */

package org.EvalGenerator;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import javax.swing.JOptionPane;

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
			
			wordMLPackage.getMainDocumentPart().addParagraphOfText(info.getInstFName());
			wordMLPackage.getMainDocumentPart().addParagraphOfText(info.getInstLName());
			wordMLPackage.getMainDocumentPart().addParagraphOfText(info.getSubject());
			
			wordMLPackage.save(new java.io.File(saveLoc.getPath() + "\\" + generateSaveString(info)));
			System.out.println("Saved to " + saveLoc.getPath() + "\\" + generateSaveString(info));
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (Exception e) {
			JOptionPane.showMessageDialog(null, "There was an error opening the file.");
		}
	}
	
	
	/** 
	 * Returns a string of a DocInfo object in the form:
	 * lname_fname subject courseNum-section semester year
	 * @param info Course info to be converted
	 * @return String of a DocInfo object in the form: lname_fname subject courseNum-section semester year
	 */
	private String generateSaveString(DocInfo info){
		return info.getInstLName()    + "_" 
				+ info.getInstFName() + " "
				+ info.getSubject()   + " "
				+ info.getCourseNum() + "-"
				+ info.getSection()   + " "
				+ info.getSemester()  + " "
				+ info.getYear() + ".docx";
	}
	
}
