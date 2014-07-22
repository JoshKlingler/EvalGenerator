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
import java.math.BigInteger;

import javax.swing.JOptionPane;

import org.docx4j.*;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;


public class WordTemplateGenerator {
	//******************* DATA MEMBERS *******************
	// Used to open files with Word after they are saved.
	private Desktop desktop;
	private WordprocessingMLPackage  wordMLPackage;
    private ObjectFactory factory;
	
	//******************* CONSTRUCTORS *******************
	public WordTemplateGenerator(){
		desktop = Desktop.getDesktop();
	}
	
	//******************* PUBLIC METHODS *******************
	/**
	 * Generates and saves Word template comments sheet based on a DocInfo object. It then opens 
	 * with Word after it is saved.
	 */
	public void generateWordTemplate(DocInfo info, File saveLoc){
		try {
			wordMLPackage = WordprocessingMLPackage.createPackage();
			factory = Context.getWmlObjectFactory();
			
			//Create table for header information
		    Tbl table = factory.createTbl();
		 
		    // Row for title
//		    Tr titleRow = factory.createTr();
//	        addStyledTableCellWithWidth(titleRow, "Student Feedback to Instructor Comments", true, "40", 20000);
	 
	        // Instructor name row
	        Tr instNameRow = factory.createTr();
	        addStyledTableCellWithWidth(instNameRow, "Instructor Name:", true, "22", 2000);
	        addStyledTableCellWithWidth(instNameRow, info.getInstFName() + " " 
	        		+ info.getInstLName(), false, "22", 6000);
	        
	        // Subject/Course number row
	        Tr courseInfoRow = factory.createTr();
	        addStyledTableCellWithWidth(courseInfoRow, "Course Subject & Number:", true, "22", 3000);
	        addStyledTableCellWithWidth(courseInfoRow, info.getSubject() + " " 
	        		+ info.getCourseNum(), false, "22", 5000);
	        
	        // Section row
	        Tr sectionRow = factory.createTr();
	        addStyledTableCellWithWidth(sectionRow, "Section:", true, "22", 2000);
	        addStyledTableCellWithWidth(sectionRow, info.getSection(), false, "22", 6000);
	        
	        
	        // Add rows to table
	        table.getContent().add(instNameRow);
	        table.getContent().add(courseInfoRow);
	        table.getContent().add(sectionRow);

	        // Add border around entire table
	        addBorders(table);
	        wordMLPackage.getMainDocumentPart().addObject(table);
		    
		    File wordDoc = createFile(saveLoc, info);
		    wordMLPackage.save(wordDoc);
		    openFile(wordDoc);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	//******************* PRIVATE METHODS *******************
	/**
     *  In this method we create a table cell,add the styling and add the cell to
     *  the table row.
     */
    private void addStyledTableCell(Tr tableRow, String content,
                        boolean bold, String fontSize) {
        Tc tableCell = factory.createTc();
        addStyling(tableCell, content, bold, fontSize);
        tableRow.getContent().add(tableCell);
    }
 
    /**
	 *  In this method we create a cell and add the given content to it.
	 *  If the given width is greater than 0, we set the width on the cell.
	 *  Finally, we add the cell to the row.
	 */
	private void addTableCellWithWidth(Tr row, String content, int width){
	    Tc tableCell = factory.createTc();
	    tableCell.getContent().add(
	        wordMLPackage.getMainDocumentPart().createParagraphOfText(
	            content));
	
	    if (width > 0) {
	        setCellWidth(tableCell, width);
	    }
	    row.getContent().add(tableCell);
	}
	
	/**
	 *  In this method we create a cell and add the given content to it.
	 *  If the given width is greater than 0, we set the width on the cell.
	 *  Finally, we add the cell to the row.
	 */
	private void addStyledTableCellWithWidth(Tr row, String content, 
			boolean bold, String fontSize, int width){
	    Tc tableCell = factory.createTc();
	    addStyling(tableCell, content, bold, fontSize);
	
	    if (width > 0) {
	        setCellWidth(tableCell, width);
	    }
	    row.getContent().add(tableCell);
	}

	/**
     *  This is where we add the actual styling information. In order to do this
     *  we first create a paragraph. Then we create a text with the content of
     *  the cell as the value. Thirdly, we create a so-called run, which is a
     *  container for one or more pieces of text having the same set of
     *  properties, and add the text to it. We then add the run to the content
     *  of the paragraph.
     *  So far what we've done still doesn't add any styling. To accomplish that,
     *  we'll create run properties and add the styling to it. These run
     *  properties are then added to the run. Finally the paragraph is added
     *  to the content of the table cell.
     */
    private void addStyling(Tc tableCell, String content,
                    boolean bold, String fontSize) {
        P paragraph = factory.createP();
 
        Text text = factory.createText();
        text.setValue(content);
 
        R run = factory.createR();
        run.getContent().add(text);
 
        paragraph.getContent().add(run);
 
        RPr runProperties = factory.createRPr();
        if (bold) {
            addBoldStyle(runProperties);
        }
 
        if (fontSize != null && !fontSize.isEmpty()) {
            setFontSize(runProperties, fontSize);
        }
 
        run.setRPr(runProperties);
 
        tableCell.getContent().add(paragraph);
    }
 
    /**
     *  In this method we're going to add the font size information to the run
     *  properties. First we'll create a half-point measurement. Then we'll
     *  set the fontSize as the value of this measurement. Finally we'll set
     *  the non-complex and complex script font sizes, sz and szCs respectively.
     */
    private void setFontSize(RPr runProperties, String fontSize) {
        HpsMeasure size = new HpsMeasure();
        size.setVal(new BigInteger(fontSize));
        runProperties.setSz(size);
        runProperties.setSzCs(size);
    }
 
    /**
     *  In this method we'll add the bold property to the run properties.
     *  BooleanDefaultTrue is the Docx4j object for the b property.
     *  Technically we wouldn't have to set the value to true, as this is
     *  the default.
     */
    private void addBoldStyle(RPr runProperties) {
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        b.setVal(true);
        runProperties.setB(b);
    }
	/**
     *  In this method we create a table cell properties object and a table width
     *  object. We set the given width on the width object and then add it to
     *  the properties object. Finally we set the properties on the table cell.
     */
    private void setCellWidth(Tc tableCell, int width) {
        TcPr tableCellProperties = new TcPr();
        TblWidth tableWidth = new TblWidth();
        tableWidth.setW(BigInteger.valueOf(width));
        tableCellProperties.setTcW(tableWidth);
        tableCell.setTcPr(tableCellProperties);
    }
	/**
	 * This method combines the save location and name of the new file to be saved. 
	 * @param savePath Folder new file is being saved in
	 * @param info Course info to be turned into a save string.
	 * @return File with name and location for saving
	 */
    private File createFile(File savePath, DocInfo info){
		return new File(savePath.getAbsolutePath() + "\\" + generateSaveFileName(info));
	}
	/**
	 * Creates and adds a cell to a table row. 
	 * @param tableRow Row to add cell to
	 * @param content String containing information for cell
	 */
	private void addTableCell(Tr tableRow, String content) {
        Tc tableCell = factory.createTc();
        tableCell.getContent().add(
        wordMLPackage.getMainDocumentPart().
            createParagraphOfText(content));
        tableRow.getContent().add(tableCell);
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
	 * Takes a DocInfo object and converts it to a string that the file
	 * will be named as.
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
	
	private void addBorders(Tbl table) {
        table.setTblPr(new TblPr());
        CTBorder border = new CTBorder();
        border.setColor("auto");
        border.setSz(new BigInteger("4"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);
 
        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setTop(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
        table.getTblPr().setTblBorders(borders);
    }
}
