/*----------------------- EvalGenerator ------------------------------
 * This program is designed to assist with the processing of student feedback forms. 
 * It automatically does most of the file management and filling in of header information,
 * allowing the user to focus solely on typing comments.
 * 
 * This class contains the code for the user interface. Most of the GUI code is auto-generated
 * using WindowBuilder. It also contains methods to handle data validation.
 */

package org.EvalGenerator;

import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.Image;
import java.awt.Toolkit;

import javax.imageio.ImageIO;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JLabel;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JButton;
import javax.swing.JComboBox;

import java.awt.Color;

import javax.swing.border.LineBorder;

import java.io.File;
import java.io.IOException;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

import javax.swing.border.TitledBorder;
import javax.swing.SwingConstants;
import javax.swing.JTabbedPane;
import javax.swing.JCheckBox;
import javax.swing.UIManager;

import org.EvalGenerator.DocInfo.Semester;


@SuppressWarnings("serial")
public class UserInterface extends JFrame {
	
	private static final String SAVE_LOC_LABEL_DEFAULT_MESSAGE = "No location selected";
	private static final int SAVE_LOCATION_LABEL_WIDTH = 182;
	private static final int WINDOW_HEIGHT = 840;
	private static final int WINDOW_WIDTH = 366;
	private static final String ICON_FILE_PATH = "files/images/logo.png";

	//******************* DATA MEMBERS *******************
	private JPanel contentPane;
	private JTextField instFNameField;
	private JTextField instLNameField;
	private JTextField subjectField;
	private JTextField courseNumField;
	private JTextField sectionField;
	private JTextField yearField;
	private JTextField facSuppNameField;
	private JTextField facSuppExtenField;
	private JTextField facSuppMailboxField;
	private JComboBox<Object> semesterComboBox;
	private JCheckBox chckbxGenerateCommentSheet;
	private JCheckBox chckbxGenerateOitScan;
	private JCheckBox chckbxSpreadsheet;
	private JFileChooser fileChooser;
	private JLabel saveLocationLabel;
	private WordTemplateGenerator wordGenerator;
	private File commentSheetSaveLoc;
	private File existingSprdshtSaveLoc;
	private File newSprdshtSaveLoc;
	private JTextField textField;
	

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UserInterface frame = new UserInterface();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	//******************* CONSTRUCTORS *******************
	/**
	 * Create the frame.
	 */
	public UserInterface() {
		// Initialize generators
		wordGenerator = new WordTemplateGenerator();
		
		// Create GUI
		setWindowSettings();
		createCourseInfoPanel();
		createOitScanInfoPanel();
		createSpreadsheetTabbedPane();
		createChkboxAndButton();
	}
	//******************* PUBLIC METHODS *******************
	//******************* PRIVATE METHODS *******************
	/**
	 * Checks for valid data based on the data appropriate for each field.
	 * Course info is always checked for validation because it is used for all 
	 * generated documents. Support staff info is only checked if an OIT 
	 * scan sheet is being generated. Save location is only checked if comment
	 * sheet is generated.
	 * 
	 * <html>Checks for:<ul>
	 *  <li>Blank fields</li>
	 *  <li>Alphabetical characters in numeric fields (Year, courseNum, extension, etc.)</li>
	 *  <li>Numeral characters in alphabetic fields (subject, etc.)</li>
	 *  <li>Length of fields</li>
	 *  <li>Valid save location</li></ul>
	 * </html>
	 * The name of any field that is invalid is added to an array that is written in a
	 * dialog box for the user if there are any messages in the array.
	 * 
	 * @param info DocInfo object containing all inputed information 
	 * @return Returns true if all data is valid and false if any data is invalid. 
	 */
	private boolean isValid(DocInfo info, boolean genOITSheet, boolean genCommentSheet){
		// Array with the name of fields with errors
		String[] errors = new String[20];
		int eIndex = 0;
		
		// Instructor first name
		if(isBlank( info.getInstFName() )){
			errors[eIndex] = "Instructor first name field blank";
			eIndex++;
		}

		// Instructor last name
		if(isBlank( info.getInstLName() )){
			errors[eIndex] = "Instructor last name field blank";
			eIndex++;
		}
		
		// Subject
		if((isBlank( info.getSubject() )) ){
			errors[eIndex] = "Subject field blank";
			eIndex++;
		}
		else if (containsNumbers( info.getSubject() )) {
			errors[eIndex] = "Subject field contains numbers";
			eIndex++;
		}
		
		// Course Number
		if(isBlank(info.getCourseNum())){
			errors[eIndex] = "Course number field blank";
			eIndex++;
		}
		else if(containsLetters( info.getCourseNum() )){
			errors[eIndex] = "Course number field contains letters";
			eIndex++;
		}
		
		// Section
		if(isBlank( info.getSection() )){
			errors[eIndex] = "Section field blank";
			eIndex++;
		}
		
		// Year
		if(isBlank( info.getYear() )){
			errors[eIndex] = "Year field blank";
			eIndex++;
		}
		else if(containsLetters( info.getYear() )){
			errors[eIndex] = "Year field contains letters";
			eIndex++;
		}

		if (genCommentSheet) {
			// Save location is only checked if the comments sheet is being generated
			if (saveLocationLabel.getText().equals(
					SAVE_LOC_LABEL_DEFAULT_MESSAGE)) {
				errors[eIndex] = "Invalid save location";
				eIndex++;
			}
		}
		// Check OIT sheet fields if OIT checkbox was checked when button was pushed
		if (genOITSheet){
			// Faculty Support Name
			if(isBlank( info.getFacSuppName() )){
				errors[eIndex] = "Faculty support name field blank";
				eIndex++;
			}
			// Support extension
			if(isBlank( info.getFacSuppExten() )){
				errors[eIndex] = "Extension field blank";
				eIndex++;
			}
			
			// Mailbox
			if(isBlank( info.getMailbox() )){
				errors[eIndex] = "Mailbox field blank";
				eIndex++;
			}
		}
		
		// If there are errors, show dialog box.
		if(eIndex > 0){
			// Put errors in bulleted list for dialog box 
			String message = "<html><b>The following fields have errors:\n</b><ul>";
			for(int i = 0; i < eIndex; i++){
				message += "<li>";
				message += errors[i];
				message += "\n</li>";
			}
			message += "</ul></html>";
			JLabel label = new JLabel(message);
			
			JOptionPane.showMessageDialog(new JFrame(), label, "Error",
			        JOptionPane.ERROR_MESSAGE);
			return false;
		}
		else{
			return true;
		}
	}
	
	/**
	 * Returns true if the given string has a length of zero.
	 */
	private boolean isBlank(String info){
		if (info.length() == 0){
			return true;
		}
		else return false;
	}
	
	/**
	 * Returns true if a string contains any letters. Otherwise returns false.
	 */
	private boolean containsLetters(String string){
		for(int i = 0; i < string.length(); i++){
			if(Character.isAlphabetic(string.charAt(i))){
				return true;
			}
		}
		return false;
	}
	
	/**
	 * Returns true if a string contains any numbers. Otherwise returns false.
	 */
	private boolean containsNumbers(String string){
		for(int i = 0; i < string.length(); i++){
			if(Character.isDigit(string.charAt(i))){
				return true;
			}
		}
		return false;
	}
	
	/**
	 * Moves data from fields into a DocInfo object.
	 */
	private DocInfo retreiveDataFromFields(){
		return new DocInfo(instFNameField.getText(), 
					instLNameField.getText(), 
					subjectField.getText(), 
					courseNumField.getText(), 
					sectionField.getText(), 
					yearField.getText(), 
					facSuppNameField.getText(),
					facSuppExtenField.getText(),
					facSuppMailboxField.getText(),
					(Semester) semesterComboBox.getSelectedItem());
	}
		
	/**
	 * Adjusts various settings of GUI such as title banner,
	 * window size, and launch location.
	 */
	private void setWindowSettings() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		// This makes the window open in the top right hand corner on launch
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		setBounds((int) screenSize.getWidth()- WINDOW_WIDTH-7 , 7, WINDOW_WIDTH, WINDOW_HEIGHT);
		
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		setResizable(false);
		
		// Change application icon from default Java icon to logo
		try {
	           Image logo = ImageIO.read(new File(ICON_FILE_PATH));
	           setIconImage(logo);
	       } catch (IOException e) {
	    	   e.printStackTrace();
	    }
		
		// Title banner at top of window
		JLabel lblTitle = new JLabel("Evaluation Template Generator");
		lblTitle.setHorizontalAlignment(SwingConstants.LEFT);
		lblTitle.setBounds(20, 13, 320, 61);
		lblTitle.setFont(new Font("Arial Bold", Font.PLAIN, 15));
		try {
	           Image img = ImageIO.read(new File(ICON_FILE_PATH));
	           lblTitle.setIcon(new ImageIcon(img));
	       } catch (IOException e) {
	    	   e.printStackTrace();
	    }
		contentPane.add(lblTitle);
	}

	/**
	 * Adds GUI elements relevant to class information section
	 */
	private void createCourseInfoPanel() {
		JPanel courseInfoPanel = new JPanel();
		courseInfoPanel.setBounds(20, 82, 320, 286);
		courseInfoPanel.setBorder(new LineBorder(new Color(0, 0, 0)));
		contentPane.add(courseInfoPanel);
		courseInfoPanel.setLayout(null);
		
		fileChooser = new JFileChooser();
		fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		
		JLabel lblInstructorFirstName = new JLabel("Instructor First Name:");
		lblInstructorFirstName.setBounds(22, 43, 126, 16);
		courseInfoPanel.add(lblInstructorFirstName);
		
		instFNameField = new JTextField();
		instFNameField.setBounds(153, 40, 128, 22);
		courseInfoPanel.add(instFNameField);
		instFNameField.setColumns(15);
		instFNameField.setText("John");//<------------------ AUTO FILL DATA FOR TEST
		
		JLabel lblInstructorLastName = new JLabel("Instructor Last Name:");
		lblInstructorLastName.setBounds(24, 70, 124, 16);
		courseInfoPanel.add(lblInstructorLastName);
		
		instLNameField = new JTextField();
		instLNameField.setBounds(153, 67, 128, 22);
		courseInfoPanel.add(instLNameField);
		instLNameField.setColumns(15);
		instLNameField.setText("Smith");//<------------------ AUTO FILL DATA FOR TEST
		
		JLabel lblSubject = new JLabel("Subject:");
		lblSubject.setBounds(100, 97, 48, 16);
		courseInfoPanel.add(lblSubject);
		
		subjectField = new JTextField();
		subjectField.setToolTipText("ex. AUT");
		subjectField.setBounds(153, 94, 48, 22);
		courseInfoPanel.add(subjectField);
		subjectField.setColumns(5);
		subjectField.setText("CST");//<------------------ AUTO FILL DATA FOR TEST
		
		JLabel lblCourseNumber = new JLabel("Course Number:");
		lblCourseNumber.setBounds(54, 124, 94, 16);
		courseInfoPanel.add(lblCourseNumber);
		
		courseNumField = new JTextField();
		courseNumField.setBounds(153, 121, 48, 22);
		courseInfoPanel.add(courseNumField);
		courseNumField.setColumns(10);
		courseNumField.setText("123");//<------------------ AUTO FILL DATA FOR TEST
		
		JLabel lblSection = new JLabel("Section:");
		lblSection.setBounds(101, 151, 47, 16);
		courseInfoPanel.add(lblSection);
		
		sectionField = new JTextField();
		sectionField.setToolTipText("ex. WN123");
		sectionField.setBounds(153, 148, 48, 22);
		courseInfoPanel.add(sectionField);
		sectionField.setColumns(10);
		sectionField.setText("WN123");//<------------------ AUTO FILL DATA FOR TEST
		
		JLabel lblSemester = new JLabel("Semester:");
		lblSemester.setBounds(88, 178, 60, 16);
		courseInfoPanel.add(lblSemester);
		
		semesterComboBox = new JComboBox<Object>(DocInfo.Semester.values());
		semesterComboBox.setBounds(153, 175, 126, 22);
		courseInfoPanel.add(semesterComboBox);
		
		JLabel lblYear = new JLabel("Year:");
		lblYear.setBounds(117, 205, 31, 16);
		courseInfoPanel.add(lblYear);
		
		yearField = new JTextField();
		yearField.setBounds(153, 202, 48, 22);
		courseInfoPanel.add(yearField);
		yearField.setColumns(10);
		yearField.setText("2014");
		
		JLabel lblCourseInformation = new JLabel("Course Information");
		lblCourseInformation.setHorizontalAlignment(SwingConstants.CENTER);
		lblCourseInformation.setFont(new Font("Arial", Font.BOLD | Font.ITALIC, 10));
		lblCourseInformation.setBounds(105, 13, 110, 16);
		courseInfoPanel.add(lblCourseInformation);
		
		JPanel saveLocPanel = new JPanel();
		saveLocPanel.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Save Location", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		saveLocPanel.setBounds(12, 230, 296, 50);
		courseInfoPanel.add(saveLocPanel);
		saveLocPanel.setLayout(null);
		
		JButton btnSaveLocBrowse = new JButton("Browse...");
		btnSaveLocBrowse.setToolTipText("Location for the template comment sheet to be saved. Generally saved in a folder specific to one semester.");
		// Open file chooser when browse button is clicked
		btnSaveLocBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				openSaveDialog();
			}
		});
		btnSaveLocBrowse.setBounds(6, 18, 91, 25);
		saveLocPanel.add(btnSaveLocBrowse);
		
		saveLocationLabel = new JLabel(SAVE_LOC_LABEL_DEFAULT_MESSAGE);
		saveLocationLabel.setBounds(102, 19, SAVE_LOCATION_LABEL_WIDTH, 22);
		saveLocPanel.add(saveLocationLabel);
	}

	/**
	 * Open save dialog and change label to name of directory if a directory was chosen
	 */
	private void openSaveDialog(){
		disableTF(fileChooser);
		int returnVal = fileChooser.showOpenDialog(UserInterface.this);
		
		if(returnVal == JFileChooser.APPROVE_OPTION){
			File file = fileChooser.getSelectedFile();
			int pathLength = file.getPath().length();
			String labelMessage;
			
			// If path length is larger than the label size, cut off
			// text to the left and replace with "..."
			if(pathLength > 25){ 
				labelMessage = "...";
				labelMessage += file.getPath().substring(pathLength - 25);
			}
			else{
				labelMessage = file.getPath();
			}
			saveLocationLabel.setText(labelMessage);
		}
	}
	
	/**
	 * Disables text field for file picker
	 */
	public boolean disableTF(Container c) {
	    Component[] cmps = c.getComponents();
	    for (Component cmp : cmps) {
	        if (cmp instanceof JTextField) {
	            ((JTextField)cmp).setEnabled(false);
	            return true;
	        }
	        if (cmp instanceof Container) {
	            if(disableTF((Container) cmp)) return true;
	        }
	    }
	    return false;
	}
	
	/** 
	 * Adds panel with fields for OIT scan sheet info to GUI.
	 */
	private void createOitScanInfoPanel() {
		JPanel oitScanInfoPanel = new JPanel();
		oitScanInfoPanel.setBounds(20, 381, 320, 134);
		oitScanInfoPanel.setBorder(new LineBorder(new Color(0, 0, 0)));
		contentPane.add(oitScanInfoPanel);
		oitScanInfoPanel.setLayout(null);
		
		JLabel lblFacultySupport = new JLabel("Faculty Support Name:");
		lblFacultySupport.setHorizontalAlignment(SwingConstants.RIGHT);
		lblFacultySupport.setBounds(12, 47, 136, 16);
		oitScanInfoPanel.add(lblFacultySupport);
		
		JLabel lblSupportExtension = new JLabel("Support Extension:");
		lblSupportExtension.setHorizontalAlignment(SwingConstants.RIGHT);
		lblSupportExtension.setBounds(40, 76, 108, 16);
		oitScanInfoPanel.add(lblSupportExtension);
		
		JLabel lblMailbox = new JLabel("Mailbox:");
		lblMailbox.setHorizontalAlignment(SwingConstants.RIGHT);
		lblMailbox.setBounds(100, 105, 48, 16);
		oitScanInfoPanel.add(lblMailbox);
		
		facSuppNameField = new JTextField();
		facSuppNameField.setBounds(153, 44, 116, 22);
		oitScanInfoPanel.add(facSuppNameField);
		facSuppNameField.setColumns(10);
		facSuppNameField.setText("Jane Doe");//<------------------ AUTO FILL DATA FOR TEST
		
		facSuppExtenField = new JTextField();
		facSuppExtenField.setBounds(153, 73, 48, 22);
		oitScanInfoPanel.add(facSuppExtenField);
		facSuppExtenField.setColumns(10);
		facSuppExtenField.setText("1234");//<------------------ AUTO FILL DATA FOR TEST
		
		facSuppMailboxField = new JTextField();
		facSuppMailboxField.setBounds(153, 102, 77, 22);
		oitScanInfoPanel.add(facSuppMailboxField);
		facSuppMailboxField.setColumns(10);
		facSuppMailboxField.setText("M-121");//<------------------ AUTO FILL DATA FOR TEST
		
		JLabel lblOitScanSheet = new JLabel("OIT Scan Sheet Information");
		lblOitScanSheet.setHorizontalAlignment(SwingConstants.CENTER);
		lblOitScanSheet.setFont(new Font("Arial", Font.BOLD | Font.ITALIC, 10));
		lblOitScanSheet.setBounds(80, 13, 160, 16);
		oitScanInfoPanel.add(lblOitScanSheet);
	}

	/**
	 * Adds tabbed view to application window to GUI to allow the user to choose between
	 * creating a new spreadsheet and using a previously generated one. 
	 */
	private void createSpreadsheetTabbedPane() {
		JTabbedPane spreadsheetTabbedPane = new JTabbedPane(JTabbedPane.TOP);
		spreadsheetTabbedPane.setBorder(new LineBorder(new Color(0, 0, 0)));
		spreadsheetTabbedPane.setBounds(20, 528, 320, 128);
		
		JPanel existSprdshtPanel = new JPanel(false);
        JLabel filler = new JLabel("test");
        filler.setHorizontalAlignment(JLabel.CENTER);
		spreadsheetTabbedPane.addTab("Existing Spreadsheet", null, existSprdshtPanel,
		                  "Use previously generated spreadsheet");
		existSprdshtPanel.setLayout(null);
		
		JPanel existSaveLocPanel = new JPanel();
		existSaveLocPanel.setLayout(null);
		existSaveLocPanel.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Save Location", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		existSaveLocPanel.setBounds(10, 8, 296, 50);
		existSprdshtPanel.add(existSaveLocPanel);
		
		JButton button = new JButton("Browse...");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// TODO File chooser
			}
		});
		button.setBounds(6, 18, 91, 25);
		existSaveLocPanel.add(button);
		
		JLabel label = new JLabel("No location selected");
		label.setBounds(102, 19, 182, 22);
		existSaveLocPanel.add(label);
		
		JPanel newSprdshtPanel = new JPanel(false);
        filler.setHorizontalAlignment(JLabel.CENTER);
		spreadsheetTabbedPane.addTab("New Spreadsheet", null, newSprdshtPanel,
		                  "Generate new spreadsheet");
		newSprdshtPanel.setLayout(null);
		
		JPanel panel = new JPanel();
		panel.setLayout(null);
		panel.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Save Location", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel.setBounds(10, 8, 296, 75);
		newSprdshtPanel.add(panel);
		
		JButton button_1 = new JButton("Browse...");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				//TODO File chooser
			}
		});
		button_1.setBounds(6, 18, 91, 25);
		panel.add(button_1);
		
		JLabel label_1 = new JLabel("No location selected");
		label_1.setBounds(102, 19, 182, 22);
		panel.add(label_1);
		
		JLabel label_2 = new JLabel("File Name:");
		label_2.setBounds(35, 47, 67, 16);
		panel.add(label_2);
		
		textField = new JTextField();
		textField.setColumns(10);
		textField.setBounds(101, 44, 171, 22);
		panel.add(textField);
		
		contentPane.add(spreadsheetTabbedPane);
	}

	/**
	 * Adds a checkbox corresponding to each of the three generated documents, as well as
	 * a button to generate documents.
	 */
	private void createChkboxAndButton() {
		chckbxGenerateCommentSheet = new JCheckBox("Generate template comment sheet");
		chckbxGenerateCommentSheet.setSelected(true);
		chckbxGenerateCommentSheet.setBounds(30, 658, 266, 25);
		contentPane.add(chckbxGenerateCommentSheet);
		
		chckbxGenerateOitScan = new JCheckBox("Generate OIT Scan Sheet");
		chckbxGenerateOitScan.setSelected(true);
		chckbxGenerateOitScan.setBounds(30, 680, 196, 25);
		contentPane.add(chckbxGenerateOitScan);
		
		chckbxSpreadsheet = new JCheckBox("Add class to spreadsheet");
		chckbxSpreadsheet.setSelected(true);
		chckbxSpreadsheet.setBounds(30, 702, 230, 25);
		contentPane.add(chckbxSpreadsheet);
		
		//---------- Action listener for main button ----------
		JButton btnGenerateDocuments = new JButton("Generate Documents");
		btnGenerateDocuments.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				DocInfo info = retreiveDataFromFields();
				
				// If data is valid in fields, generate documents that have checked boxes 
				if(isValid(info, chckbxGenerateOitScan.isSelected(), 
						chckbxGenerateCommentSheet.isSelected() )){
					generateDocuments(info);
				}
			}
		});
		btnGenerateDocuments.setBounds(99, 732, 161, 45);
		contentPane.add(btnGenerateDocuments);
	}
	/**
	 * Generates documents based on which checkboxes have been checked on the GUI.
	 * @param info Data used to generate documents
	 */
	private void generateDocuments(DocInfo info){
		contentPane.setCursor(new Cursor(Cursor.WAIT_CURSOR));
		
		if( chckbxGenerateCommentSheet.isSelected() ){
			wordGenerator.generateCommentTemplate(info, fileChooser.getSelectedFile());
			System.out.println("Word");
		}
		if( chckbxGenerateOitScan.isSelected() ){
			wordGenerator.generateOITSheet(info);
		}
		if( chckbxSpreadsheet.isSelected() ){
			//TODO Spreadsheet generator
			System.out.println("SS");
		}
		
		contentPane.setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
	}
}
