/**
 * This class is responsible for generating the template Word
 * documents. It uses Docx4j, a Java library designed to create
 * .docx documents. It also opens the files after they are created using
 * the Desktop class, which is part of the default JDK.
 * 
 * Some of the table generation code was taken from:
 * http://blog.iprofs.nl/2012/09/06/creating-word-documents-with-docx4j/
 */

package org.EvalGenerator;

import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
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
	private static final int TABLE_SHORT_FIELD_LENGTH = 2000;
	private static final String OIT_SCAN_SHEET_SAVE_LOC = "files/textDocs/OITScanSheet.docx";
	private static final String COMMENT_SHEET_QUESTION_SAVE_PATH = "files/textDocs/evalQuestions.txt";
	private static final int TABLE_LONG_FIELD_LENGTH = 6250;
	private static final String NEW_LINE = "\n\r";
	//******************* DATA MEMBERS *******************
	// Used to open files with Word after they are saved.
	private Desktop desktop;
	
	// Used for Word doc creation
	private WordprocessingMLPackage  wordMLPackage;
    private ObjectFactory factory;
    
    // Stores questions for comments sheet that are read from a file.
    private ArrayList<String> evalQuestions = new ArrayList<String>();
	
	//******************* CONSTRUCTORS *******************
	public WordTemplateGenerator(){
		// Read and store course evaluation questions
		readQuestionsFromFile();
		
		// Used to open files
		desktop = Desktop.getDesktop();
	}
	
	//******************* PUBLIC METHODS *******************
	/**
	 * Generates, saves, and opens Word template comments sheet based on a DocInfo object. 
	 * It generates a header table with course info stored in the DocInfo object. Each
	 * evaluation question has its own table below the header table. The file is saved at
	 * the location specified with a name generated by the program based on the course info.
	 */
	public void generateCommentTemplate(DocInfo info, File saveLoc){
		try {
			// Check if file exists. If it does exist, give the user the option
			// to cancel. 
			File wordDoc = new File(saveLoc, generateSaveFileName(info));
			if( wordDoc.exists() ){
				String message = "The comment template sheet already exists. By selecting "
						+ "yes, the file will be overwritten. Are you sure you want to overwrite "
						+ "the file?";
				int response = JOptionPane.showConfirmDialog(null, message);
				if(response != JOptionPane.YES_OPTION){
					return;
				}
			}
			
			wordMLPackage = WordprocessingMLPackage.createPackage();
			factory = Context.getWmlObjectFactory();
			
			// Add title to top of document
			wordMLPackage.getMainDocumentPart().addParagraphOfText("Student Feedback to Instructor");

			//Create table for header information
		    Tbl courseInfoTable = factory.createTbl();
	 
	        // Create instructor name row
	        Tr instNameRow = factory.createTr();
	        addStyledTableCellWithWidth(instNameRow, "Instructor Name:", true, "20", TABLE_SHORT_FIELD_LENGTH);
	        addStyledTableCellWithWidth(instNameRow, info.getInstFName() + " " 
	        		+ info.getInstLName(), false, "20", TABLE_LONG_FIELD_LENGTH);
	        
	        // Create subject/course number row
	        Tr courseInfoRow = factory.createTr();
	        addStyledTableCellWithWidth(courseInfoRow, "Course Subject & Number:", true, "20", 3000);
	        addStyledTableCellWithWidth(courseInfoRow, info.getSubject() + " " 
	        		+ info.getCourseNum(), false, "20", TABLE_LONG_FIELD_LENGTH);
	        
	        // Create section row
	        Tr sectionRow = factory.createTr();
	        addStyledTableCellWithWidth(sectionRow, "Section:", true, "20", TABLE_SHORT_FIELD_LENGTH);
	        addStyledTableCellWithWidth(sectionRow, info.getSection(), false, "20", TABLE_LONG_FIELD_LENGTH);
	        
	        // Semester/Year row
	        Tr semesterRow = factory.createTr();
	        addStyledTableCellWithWidth(semesterRow, "Semester:", true, "20", TABLE_SHORT_FIELD_LENGTH);
	        addStyledTableCellWithWidth(semesterRow, info.getSemester().toString() 
	        		+ " " + info.getYear(), false, "20", TABLE_LONG_FIELD_LENGTH);
	        
	        // Add rows to table
	        courseInfoTable.getContent().add(instNameRow);
	        courseInfoTable.getContent().add(courseInfoRow);
	        courseInfoTable.getContent().add(sectionRow);
	        courseInfoTable.getContent().add(semesterRow);
	        
	        // Add border around entire table
	        addBorders(courseInfoTable);
	        
	        // Place tables on document
	        wordMLPackage.getMainDocumentPart().addObject(courseInfoTable);
	        
	        // Create a new table for each evaluation question
	        for(String question:evalQuestions){
	        	// Add empty paragraph so there is a space between questions
	        	wordMLPackage.getMainDocumentPart().addParagraphOfText(NEW_LINE);
	        	
	        	// Create table and add to doc 
	        	Tbl questionTable = factory.createTbl();
	        	
	        	// Row with question
	        	Tr questionRow = factory.createTr();
	        	addStyledTableCellWithWidth(questionRow, question, true, "20", 10000);
	        	questionTable.getContent().add(questionRow);
	        	
	        	// Row with blank line for input
	        	Tr blankRow = factory.createTr();
	        	addStyledTableCellWithWidth(blankRow, "", true, "20", 10000);
	        	questionTable.getContent().add(blankRow); 
	        	
	        	addBorders(questionTable);
	        	
	        	// Add table to document
	        	wordMLPackage.getMainDocumentPart().addObject(questionTable);
	        }
	        
		    
		    wordMLPackage.save(wordDoc);
		    openFile(wordDoc);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Docx4JException e) {
			JOptionPane.showMessageDialog(null, "The comment sheet could not be saved. Make sure the document is not already open in Word.");
			e.printStackTrace();
		}
	}
	
	
	/**
	 * Generates an OIT scan sheet. If checkbox for print is open, uses a Desktop
	 * object to print to the default printer. The current system date is used as the date. 
	 * @param info Course inform
	 * @param print If true, the document will be sent to default printer
	 */
	public void generateOITSheet(DocInfo info, boolean print){
		try {		
			// Create new doc
			wordMLPackage = WordprocessingMLPackage.createPackage();
			factory = Context.getWmlObjectFactory();
			
			// Add title to top of document
			wordMLPackage.getMainDocumentPart().addParagraphOfText("OIT Scan Cover Sheet");

			//Create table for header information
			Tbl infoTable = factory.createTbl();

			// Create instructor name row
			Tr instNameRow = factory.createTr();
			addStyledTableCellWithWidth(instNameRow, "Instructor Name:", true, "20", TABLE_SHORT_FIELD_LENGTH);
			addStyledTableCellWithWidth(instNameRow, info.getInstFName() + " " 
					+ info.getInstLName(), false, "20", TABLE_LONG_FIELD_LENGTH);

			// Create subject/course number row
			Tr courseInfoRow = factory.createTr();
			addStyledTableCellWithWidth(courseInfoRow, "Course Subject & Number:", true, "20", 3000);
			addStyledTableCellWithWidth(courseInfoRow, info.getSubject() + " " 
					+ info.getCourseNum(), false, "20", TABLE_LONG_FIELD_LENGTH);

			// Create section row
			Tr sectionRow = factory.createTr();
			addStyledTableCellWithWidth(sectionRow, "Section:", true, "20", TABLE_SHORT_FIELD_LENGTH);
			addStyledTableCellWithWidth(sectionRow, info.getSection(), false, "20", TABLE_LONG_FIELD_LENGTH);

			// Semester/Year row
			Tr semesterRow = factory.createTr();
			addStyledTableCellWithWidth(semesterRow, "Semester:", true, "20", TABLE_SHORT_FIELD_LENGTH);
			addStyledTableCellWithWidth(semesterRow, info.getSemester().toString() 
					+ " " + info.getYear(), false, "20", TABLE_LONG_FIELD_LENGTH);
			
			// Faculty supp name row
			Tr facSuppNameRow = factory.createTr();
			addStyledTableCellWithWidth(facSuppNameRow, "Faculty Support Name:", true, "20", TABLE_SHORT_FIELD_LENGTH);
			addStyledTableCellWithWidth(facSuppNameRow, info.getFacSuppName(), false, "20", TABLE_LONG_FIELD_LENGTH);

			// Faculty supp name row
			Tr facSuppExtenRow = factory.createTr();
			addStyledTableCellWithWidth(facSuppExtenRow, "Faculty Support Extension:", true, "20", TABLE_SHORT_FIELD_LENGTH);
			addStyledTableCellWithWidth(facSuppExtenRow, info.getFacSuppExten(), false, "20", TABLE_LONG_FIELD_LENGTH);
			
			// Mailbox row
			Tr mailboxRow = factory.createTr();
			addStyledTableCellWithWidth(mailboxRow, "I would like the results delivered to mailbox:", true, "20", TABLE_SHORT_FIELD_LENGTH);
			addStyledTableCellWithWidth(mailboxRow, info.getMailbox(), false, "20", TABLE_LONG_FIELD_LENGTH);
			
			Tr dateRow = factory.createTr();
			DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
			Date date = new Date();
			System.out.println(dateFormat.format(date));
			addStyledTableCellWithWidth(dateRow, "Date of Request:", true, "20", TABLE_SHORT_FIELD_LENGTH);
			addStyledTableCellWithWidth(dateRow, dateFormat.format(date), false, "20", TABLE_LONG_FIELD_LENGTH);
			
			// Add rows to table
			infoTable.getContent().add(instNameRow);
			infoTable.getContent().add(courseInfoRow);
			infoTable.getContent().add(sectionRow);
			infoTable.getContent().add(semesterRow);
			infoTable.getContent().add(facSuppNameRow);
			infoTable.getContent().add(facSuppExtenRow);
			infoTable.getContent().add(mailboxRow);
			infoTable.getContent().add(dateRow);

			// Add border around entire table
			addBorders(infoTable);

			// Place tables on document
			wordMLPackage.getMainDocumentPart().addObject(infoTable);
			
			// Save in project files
			wordMLPackage.save(new File(OIT_SCAN_SHEET_SAVE_LOC));
			
			// Print or open file based on user decision
			if(print){
				desktop.print(new File(OIT_SCAN_SHEET_SAVE_LOC));
			}
			else{
				desktop.open(new File(OIT_SCAN_SHEET_SAVE_LOC));
			}
			
			
			System.out.println("OIT Scan sheet generated.");
			
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			JOptionPane.showMessageDialog(null, "Unable to open the OIT scan sheet. Make sure the document is not open.");
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//******************* PRIVATE METHODS *******************
	/**
	 * Read questions for comment sheet from file and store.
	 */
	private void readQuestionsFromFile(){
		try {
			// Open file
			BufferedReader reader = new BufferedReader(new FileReader(new File(COMMENT_SHEET_QUESTION_SAVE_PATH)));
			
			// Loop until all lines of the file are stored in the array.
			while(reader.ready()){
				evalQuestions.add(reader.readLine());
			}
			
			reader.close();
						
		} catch (FileNotFoundException e) {
			JOptionPane.showMessageDialog(new JFrame(), "Unable to open questions for evaluation sheet.");
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
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
