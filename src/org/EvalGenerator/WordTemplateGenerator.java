/**
 * This class is responsible for generating the template comments
 * sheets that the comments are written in. 
 */

package org.EvalGenerator;

import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

import javax.swing.JOptionPane;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;


public class WordTemplateGenerator {
	
	// Used to open files with Word after they are saved.
	private Desktop desktop;
	
	public WordTemplateGenerator(){
		desktop = Desktop.getDesktop();
	}
	
	/**
	 * Generates and saves Word template comments sheet based on a DocInfo object. It then opens 
	 * with Word after it is saved.
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
			wordMLPackage.getMainDocumentPart().addParagraphOfText(info.getCourseNum());
			wordMLPackage.getMainDocumentPart().addParagraphOfText(info.getSection());
			
			// Save to designated location with file name generated using course info
			File wordDoc = new java.io.File(saveLoc.getPath() + "\\" + generateSaveFileName(info));
			wordMLPackage.save(wordDoc);
			openFile(wordDoc);
			
			System.out.println("Saved to " + saveLoc.getPath() + "\\" + generateSaveFileName(info));
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (Docx4JException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(null, "There was an error saving the file.");
		}
	}
	
	/**
	 * Opens designated file with default program 
	 * @param file File to be opened
	 */
	private void openFile(File file) {      
		try {
			desktop.open(file);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/** 
	 * Returns a string of a DocInfo object in the form:
	 * lname_fname subject courseNum-section semester year
	 * @param info Course info to be converted
	 * @return String of a DocInfo object in the form: lname_fname subject courseNum-section semester year
	 */
	private String generateSaveFileName(DocInfo info){
		return info.getInstLName()    + "_" 
				+ info.getInstFName() + " "
				+ info.getSubject()   + " "
				+ info.getCourseNum() + "-"
				+ info.getSection()   + " "
				+ info.getSemester()  + " "
				+ info.getYear() + ".docx";
	}
	
}
