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
import javax.swing.SwingUtilities;

import java.awt.Color;

import javax.swing.border.LineBorder;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.prefs.Preferences;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;

import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.SwingConstants;
import javax.swing.JTabbedPane;
import javax.swing.JCheckBox;
import javax.swing.UIManager;

import org.EvalGenerator.DocInfo.Semester;


@SuppressWarnings("serial")
public class UserInterface extends JFrame {
	
	private static final String EXIST_SPRDSHT_LBL_DEFAULT_MESSAGE = "No file selected";
	private static final String NEW_SPRDSHT_LBL_DEFAULT_MESSAGE = "No location selected";
	private static final int SPREADSHEET_NONE = 0;
	private static final int SPREADSHEET_NEW = 2;
	private static final int SPREADSHEET_EXISTING = 1;
	private static final int EXIST_SPRDSHT_INDEX = 0;
	private static final String SAVE_LOC_LABEL_DEFAULT_MESSAGE = NEW_SPRDSHT_LBL_DEFAULT_MESSAGE;
	private static final int SAVE_LOCATION_LABEL_WIDTH = 182;
	private static final int WINDOW_HEIGHT = 900;
	private static final int WINDOW_WIDTH = 366;
	private static final String ICON_FILE_PATH = "files/images/logo.png";

	//******************* DATA MEMBERS *******************
	private JPanel contentPane;
	private JPanel oitScanInfoPanel = new JPanel();
	private JPanel existSprdshtPanel = new JPanel(false);
	private JPanel newSprdshtPanel = new JPanel(false);
	private JPanel spreadsheetPanel = new JPanel();
	
	private JTextField fldInstFName;
	private JTextField fldInstLName;
	private JTextField fldSubject;
	private JTextField fldCourseNum;
	private JTextField fldSection;
	private JTextField fldYear;
	private JTextField fldFacSuppName;
	private JTextField fldFacSuppExten;
	private JTextField fldFacSuppMailbox;
	private JTextField fldNewSprdshtFileName;
	
	private JComboBox<Object> semesterComboBox;
	
	private JCheckBox chckbxGenerateCommentSheet;
	private JCheckBox chckbxGenerateOitScan;
	private JCheckBox chckbxSpreadsheet;
	private JCheckBox chckboxPrintOITSheet;
	
	private JFileChooser commentSheetFileChooser = new JFileChooser();
	private JFileChooser existSprdshtFileChooser = new JFileChooser();
	private JFileChooser newSprdshtFileChooser   = new JFileChooser();
	
	private JLabel lblCommentSaveLoc;
	private JLabel lblExistSprdshtSaveLoc;
	private JLabel lblNewSprdshtSaveLoc;
	
	private WordTemplateGenerator wordGenerator = new WordTemplateGenerator();
	private SpreadsheetManager sprdshtManager = new SpreadsheetManager();
		
	private JTabbedPane spreadsheetTabbedPane;
	
	private Dimension screenSize;
	


	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		
		// Change look and feel of application
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		}catch(Exception ex) {
			ex.printStackTrace();
		}
		
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
		// Store size of screen
		screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		
		// Change look and feel of application
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		}catch(Exception ex) {
			ex.printStackTrace();
		}
		
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
	 * Generates documents based on which checkboxes have been checked on the GUI. Assumes
	 * all data validation has already been done. 
	 * @param info Data used to generate documents
	 */
	private void generateDocuments(DocInfo info){
		
		// Spinning cursor to show that work is being done
		contentPane.setCursor(new Cursor(Cursor.WAIT_CURSOR));
		
		// Eval comment sheet
		if( chckbxGenerateCommentSheet.isSelected() ){
			wordGenerator.generateCommentTemplate(info, commentSheetFileChooser.getSelectedFile());
		}
		// Oit Scan Sheet
		if( chckbxGenerateOitScan.isSelected() ){
			wordGenerator.generateOITSheet(info, chckboxPrintOITSheet.isSelected());
		}
		// Spreadsheet
		if( chckbxSpreadsheet.isSelected() ){
			// Determine whether or not we are generating a new spreadsheet 
			// or using an old one based on which tab of the tabbed pane is selected.
			int index = spreadsheetTabbedPane.getSelectedIndex();

			// Use existing CSV file
			if(index == EXIST_SPRDSHT_INDEX){
				File saveLoc = existSprdshtFileChooser.getSelectedFile();
				sprdshtManager.addClassToSpreadsheet(saveLoc, info);
			}
			// Create new CSV file. Confirmation dialog is shown when this is chosen
			else{
				// Dialog box
				String message = "Are you sure you want to create a new spreadsheet?";
				int choice = JOptionPane.showOptionDialog(new JFrame(), message, "Confirmation", JOptionPane.YES_NO_OPTION, 
						JOptionPane.QUESTION_MESSAGE, null, null, null);
				
				// If user clicks yes
				if(choice == JOptionPane.YES_OPTION){
					File saveLoc = newSprdshtFileChooser.getSelectedFile();
					sprdshtManager.createNewSpreadsheet(saveLoc, fldNewSprdshtFileName.getText(), info);
				}				
			}
		}
		
		// Switch back to normal cursor
		contentPane.setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
	}

	/**
	 * Enables or disables all components in panel. If one of the components
	 * is a JPanel, it recursively calls the function and disables that too.
	 * @param panel Panel with components to have enabled state changed
	 * @param enable If true, the components are to be enabled. If false, they
	 * will be disabled. 
	 */
	private void changePanelDisableState(JPanel panel, Boolean enable){
		Component[] components = panel.getComponents();
	
		for(Component c:components){
			c.setEnabled(enable);
			if(c.getClass() == JPanel.class){
				changePanelDisableState((JPanel) c, enable);
			}
			if(c.getClass() == JTabbedPane.class){
				changeTabbedPanelDisableState((JTabbedPane) c, enable);
			}
		}
	}
	
	/**
	 * Disables the tabbed pane and calls changePanelDisableState for all
	 * panels attached to the pane.
	 * @param pane JTabbedPane to be disabled
	 * @param enable If true, the components are to be enabled. If false, they
	 * will be disabled.
	 */
	private void changeTabbedPanelDisableState(JTabbedPane pane, Boolean enable){
		Component[] components = pane.getComponents();
	
		for(Component c:components){
			
			if(c.getClass() == JPanel.class){
				changePanelDisableState((JPanel) c, enable);
			}
			
			c.setEnabled(enable);
		}

		pane.setEnabled(enable);
	}
	
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
	 * @param genOITSheet If true, OIT sheet is being generated so relevant information is checked.
	 * @param genOITSheet If true, comment sheet is being generated so relevant information is checked.
	 * @return Returns true if all data is valid and false if any data is invalid. 
	 */
	private boolean isValid(DocInfo info, boolean genOITSheet, boolean genCommentSheet, int spreadsheetChoice){		
		ArrayList<String> error = new ArrayList<String>();
		
		// Instructor first name
		if( info.getInstFName().isEmpty() ){
			error.add("Instructor first name field blank");
		}

		// Instructor last name
		if( info.getInstLName().isEmpty() ){
			error.add("Instructor last name field blank");
		}
		
		// Subject
		info.setSubject(info.getSubject().toUpperCase());
		if( info.getSubject().isEmpty() ){
			error.add("Subject field blank");
		}
		else if (containsNumbers( info.getSubject() )) {
			error.add("Subject field contains numbers");
		}
		
		// Course Number
		if( info.getCourseNum().isEmpty() ){
			error.add("Course number field blank");
		}
		else if(containsLetters( info.getCourseNum() )){
			error.add("Course number field contains letters");
		}
		
		// Section
		info.setSection(info.getSection().toUpperCase());
		if( info.getSection().isEmpty() ){
			error.add("Section field blank");
		}
		
		// Year
		if( info.getYear().isEmpty() ){
			error.add("Year field blank");
		}
		else if(containsLetters( info.getYear() )){
			error.add("Year field contains letters");
		}

		if (genCommentSheet) {
			// Save location is only checked if the comments sheet is being generated
			if (lblCommentSaveLoc.getText().equals(
					SAVE_LOC_LABEL_DEFAULT_MESSAGE)) {
				error.add("Invalid comment sheet save location");
			}
		}
		// Check OIT sheet fields if OIT checkbox was checked when button was pushed
		if (genOITSheet){
			// Faculty Support Name
			if( info.getFacSuppName().isEmpty() ){
				error.add("Faculty support name field blank");
			}
			// Support extension
			if( info.getFacSuppExten().isEmpty() ){
				error.add("Extension field blank");
			}
			
			// Mailbox
			if( info.getMailbox().isEmpty() ){
				error.add("Mailbox field blank");
			}
		}
		
		switch(spreadsheetChoice){
		case(SPREADSHEET_NONE):	// Do nothing
								break;
		
		case(SPREADSHEET_NEW):	// If save location has not been chosen at all
								if(lblNewSprdshtSaveLoc.getText().equals(NEW_SPRDSHT_LBL_DEFAULT_MESSAGE)){
									error.add("A location must be selected to save the new spreadsheet");
								}
								// If location has been chosen
								else{
									// Check if name is blank
									if(fldNewSprdshtFileName.getText().length() == 0){
										error.add("New spreadsheet name field blank");
									}
									// Check if file exists
									else if(new File(newSprdshtFileChooser.getSelectedFile(), 
											fldNewSprdshtFileName.getText() + ".csv").exists()){
										error.add("Spreadsheet file already exists");
									}	
								}
								break;
								
		case(SPREADSHEET_EXISTING): // Check if file has been chosen
								if(lblExistSprdshtSaveLoc.getText().equals(EXIST_SPRDSHT_LBL_DEFAULT_MESSAGE)){
									error.add("Spreadsheet must be chosen");
								}
								else{ // File has been chosen. Check if it is a .csv file
									String path = existSprdshtFileChooser.getSelectedFile().getAbsolutePath();
									if(!path.contains(".csv")){
										error.add("The spreadsheet must be a .csv file");
									}
								}
		}
		
		// If there are errors, show dialog box.
		if(!error.isEmpty()){
			// Put errors in an html bulleted list for dialog box 
			String message = "<html><b>The following fields have errors:\n</b><ul>";
			for(String e:error){
				message += "<li>";
				message += e;
				message += "\n</li>";
			}
			message += "</ul></html>";
			JLabel label = new JLabel(message);
			
			JOptionPane.showMessageDialog(new JFrame(), label, "Error",
			        JOptionPane.ERROR_MESSAGE);
			return false;
		}
		else{// All data is valid
			return true;
		}
	}
	
	/**
	 * Open save dialog and change label to name of directory if a directory was chosen
	 * @param saveLoc Stores location for file to be saved to.
	 * @param label Label that will change to show the selected file path.
	 */
	private void openSaveDialog(JFileChooser fileChooser, JLabel label, boolean directoryOnly){
		
		
		
		disableTF(fileChooser);
		
		// If selecting a save location for either the comment sheet or new spreadsheet,
		// the user can only select a folder. If selecting an existing spreadsheet they 
		// can only select the CSV file.
		if(directoryOnly){
			fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		}
		else{
			fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
			FileNameExtensionFilter filter = new FileNameExtensionFilter("CSV Files", "csv");
			fileChooser.setFileFilter(filter);
		}
		
		// Open file dialog window
		int returnVal = fileChooser.showOpenDialog(this);
		
		if(returnVal == JFileChooser.APPROVE_OPTION){
			File saveLoc = fileChooser.getSelectedFile();
			int pathLength = saveLoc.getPath().length();
			String labelMessage;
			
			// If path is too large to fit, cut off
			// text to the left and replace with "..."
			if(pathLength > 25){ 
				labelMessage = "...";
				labelMessage += saveLoc.getPath().substring(pathLength - 25);
			}
			else{
				labelMessage = saveLoc.getPath();
			}
			label.setText(labelMessage);
		}
	}

	/**
	 * Moves data from fields into a DocInfo object.
	 */
	private DocInfo retreiveDataFromFields(){
		return new DocInfo(fldInstFName.getText(), 
					fldInstLName.getText(), 
					fldSubject.getText(), 
					fldCourseNum.getText(), 
					fldSection.getText(), 
					fldYear.getText(), 
					fldFacSuppName.getText(),
					fldFacSuppExten.getText(),
					fldFacSuppMailbox.getText(),
					(Semester) semesterComboBox.getSelectedItem());
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
	 * Adjusts various settings of GUI such as title banner,
	 * window size, and launch location.
	 */
	private void setWindowSettings() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		
		
		// This makes the window open in the top right hand corner on launch
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
		try {
			Image img = ImageIO.read(new File(ICON_FILE_PATH));
			// Title banner at top of window
			JLabel lblTitle = new JLabel("Evaluation Template Generator");
			lblTitle.setHorizontalAlignment(SwingConstants.LEFT);
			lblTitle.setBounds(20, 13, 320, 61);
			lblTitle.setFont(new Font("Arial Bold", Font.PLAIN, 15));
	   		lblTitle.setIcon(new ImageIcon(img));
	   		contentPane.add(lblTitle);
	       } catch (IOException e) {
	    	   e.printStackTrace();
	    }
		
		
	}

	/**
	 * Adds GUI elements relevant to class information section
	 */
	private void createCourseInfoPanel() {
		JPanel courseInfoPanel = new JPanel();
		courseInfoPanel.setBounds(20, 145, 320, 286);
		courseInfoPanel.setBorder(new LineBorder(new Color(0, 0, 0)));
		contentPane.add(courseInfoPanel);
		courseInfoPanel.setLayout(null);
		
		commentSheetFileChooser = new JFileChooser();
		
		JLabel lblInstructorFirstName = new JLabel("Instructor First Name:");
		lblInstructorFirstName.setBounds(22, 43, 126, 16);
		courseInfoPanel.add(lblInstructorFirstName);
		
		fldInstFName = new JTextField();
		fldInstFName.setBounds(153, 40, 128, 22);
		courseInfoPanel.add(fldInstFName);
		fldInstFName.setColumns(15);
		fldInstFName.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		                fldInstFName.selectAll();
		            }
		        });
		    }
		});
		
		JLabel lblInstructorLastName = new JLabel("Instructor Last Name:");
		lblInstructorLastName.setBounds(24, 70, 124, 16);
		courseInfoPanel.add(lblInstructorLastName);
		
		fldInstLName = new JTextField();
		fldInstLName.setBounds(153, 67, 128, 22);
		courseInfoPanel.add(fldInstLName);
		fldInstLName.setColumns(15);
		fldInstLName.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		                fldInstLName.selectAll();
		            }
		        });
		    }
		});
		
		
		JLabel lblSubject = new JLabel("Subject:");
		lblSubject.setBounds(100, 97, 48, 16);
		courseInfoPanel.add(lblSubject);
		
		fldSubject = new JTextField();
		fldSubject.setToolTipText("ex. AUT");
		fldSubject.setBounds(153, 94, 48, 22);
		courseInfoPanel.add(fldSubject);
		fldSubject.setColumns(5);
		fldSubject.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		            	fldSubject.selectAll();
		            }
		        });
		    }
		});
		
		JLabel lblCourseNumber = new JLabel("Course Number:");
		lblCourseNumber.setBounds(54, 124, 94, 16);
		courseInfoPanel.add(lblCourseNumber);
		
		fldCourseNum = new JTextField();
		fldCourseNum.setBounds(153, 121, 48, 22);
		courseInfoPanel.add(fldCourseNum);
		fldCourseNum.setColumns(10);
		fldCourseNum.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		            	fldCourseNum.selectAll();
		            }
		        });
		    }
		});
		
		JLabel lblSection = new JLabel("Section:");
		lblSection.setBounds(101, 151, 47, 16);
		courseInfoPanel.add(lblSection);
		
		fldSection = new JTextField();
		fldSection.setToolTipText("ex. WN123");
		fldSection.setBounds(153, 148, 48, 22);
		courseInfoPanel.add(fldSection);
		fldSection.setColumns(10);
		fldSection.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		            	fldSection.selectAll();
		            }
		        });
		    }
		});
		
		JLabel lblSemester = new JLabel("Semester:");
		lblSemester.setBounds(88, 178, 60, 16);
		courseInfoPanel.add(lblSemester);
		
		semesterComboBox = new JComboBox<Object>(DocInfo.Semester.values());
		semesterComboBox.setBounds(153, 175, 126, 22);
		courseInfoPanel.add(semesterComboBox);
		
		JLabel lblYear = new JLabel("Year:");
		lblYear.setBounds(117, 205, 31, 16);
		courseInfoPanel.add(lblYear);
		
		fldYear = new JTextField();
		fldYear.setBounds(153, 202, 48, 22);
		courseInfoPanel.add(fldYear);
		fldYear.setColumns(10);
		fldYear.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		            	fldYear.selectAll();
		            }
		        });
		    }
		});
		
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
		
		JButton btnCommentSaveLocBrowse = new JButton("Browse...");
		btnCommentSaveLocBrowse.setToolTipText("Location for the template comment sheet to be saved. Generally saved in a folder specific to one semester.");
		// Open file chooser when browse button is clicked
		btnCommentSaveLocBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				openSaveDialog(commentSheetFileChooser, lblCommentSaveLoc, true);
			}
		});
		btnCommentSaveLocBrowse.setBounds(6, 18, 91, 25);
		saveLocPanel.add(btnCommentSaveLocBrowse);
		
		lblCommentSaveLoc = new JLabel(SAVE_LOC_LABEL_DEFAULT_MESSAGE);
		lblCommentSaveLoc.setBounds(102, 19, SAVE_LOCATION_LABEL_WIDTH, 22);
		saveLocPanel.add(lblCommentSaveLoc);
	}

	/**
	 * Disables text field for file picker. 
	 */
	private boolean disableTF(Container c) {
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
		oitScanInfoPanel.setBounds(20, 444, 320, 156);
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
		
		fldFacSuppName = new JTextField();
		fldFacSuppName.setBounds(153, 44, 116, 22);
		oitScanInfoPanel.add(fldFacSuppName);
		fldFacSuppName.setColumns(10);
		fldFacSuppName.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		            	fldFacSuppName.selectAll();
		            }
		        });
		    }
		});
		
		fldFacSuppExten = new JTextField();
		fldFacSuppExten.setBounds(153, 73, 48, 22);
		oitScanInfoPanel.add(fldFacSuppExten);
		fldFacSuppExten.setColumns(10);
		fldFacSuppExten.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		            	fldFacSuppExten.selectAll();
		            }
		        });
		    }
		});
		
		fldFacSuppMailbox = new JTextField();
		fldFacSuppMailbox.setBounds(153, 102, 77, 22);
		oitScanInfoPanel.add(fldFacSuppMailbox);
		fldFacSuppMailbox.setColumns(10);
		fldFacSuppMailbox.addFocusListener(new java.awt.event.FocusAdapter() {
		    public void focusGained(java.awt.event.FocusEvent evt) {
		        SwingUtilities.invokeLater(new Runnable() {
		            @Override
		            public void run() {
		            	fldFacSuppMailbox.selectAll();
		            }
		        });
		    }
		});
		
		JLabel lblOitScanSheet = new JLabel("OIT Scan Sheet Information");
		lblOitScanSheet.setHorizontalAlignment(SwingConstants.CENTER);
		lblOitScanSheet.setFont(new Font("Arial", Font.BOLD | Font.ITALIC, 10));
		lblOitScanSheet.setBounds(80, 13, 160, 16);
		oitScanInfoPanel.add(lblOitScanSheet);
		
		chckboxPrintOITSheet = new JCheckBox("Print document to default printer");
		chckboxPrintOITSheet.setSelected(true);
		chckboxPrintOITSheet.setBounds(53, 130, 215, 25);
		oitScanInfoPanel.add(chckboxPrintOITSheet);
		
	}

	/**
	 * Adds tabbed view to application window to GUI to allow the user to choose between
	 * creating a new spreadsheet and using a previously generated one. 
	 */
	private void createSpreadsheetTabbedPane() {
		
		spreadsheetPanel.setBorder(new LineBorder(new Color(0, 0, 0)));
		spreadsheetPanel.setBounds(20, 614, 320, 170);
		contentPane.add(spreadsheetPanel);
		spreadsheetPanel.setLayout(null);
		spreadsheetTabbedPane = new JTabbedPane(JTabbedPane.TOP);
		spreadsheetTabbedPane.setBounds(3, 40, 316, 128);
		spreadsheetPanel.add(spreadsheetTabbedPane);
		spreadsheetTabbedPane.setBorder(null);
		
		spreadsheetTabbedPane.addTab("Existing Spreadsheet", null, existSprdshtPanel,
		                  "Use previously generated spreadsheet");
		existSprdshtPanel.setLayout(null);
		
		JPanel existSaveLocPanel = new JPanel();
		existSaveLocPanel.setLayout(null);
		existSaveLocPanel.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "File Location", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		existSaveLocPanel.setBounds(10, 8, 296, 50);
		existSprdshtPanel.add(existSaveLocPanel);
		
		JButton btnExistSprdshtSaveLoc = new JButton("Browse...");
		btnExistSprdshtSaveLoc.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openSaveDialog(existSprdshtFileChooser, lblExistSprdshtSaveLoc, false);
			}
		});
		btnExistSprdshtSaveLoc.setBounds(6, 18, 91, 25);
		existSaveLocPanel.add(btnExistSprdshtSaveLoc);
		
		lblExistSprdshtSaveLoc = new JLabel(EXIST_SPRDSHT_LBL_DEFAULT_MESSAGE);
		lblExistSprdshtSaveLoc.setBounds(102, 19, 182, 22);
		existSaveLocPanel.add(lblExistSprdshtSaveLoc);
		

		spreadsheetTabbedPane.addTab("New Spreadsheet", null, newSprdshtPanel,
		                  "Generate new spreadsheet");
		newSprdshtPanel.setLayout(null);
		
		JPanel panel = new JPanel();
		panel.setLayout(null);
		panel.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Save Location", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel.setBounds(10, 8, 296, 75);
		newSprdshtPanel.add(panel);
		
		lblNewSprdshtSaveLoc = new JLabel(NEW_SPRDSHT_LBL_DEFAULT_MESSAGE);
		lblNewSprdshtSaveLoc.setBounds(102, 19, 182, 22);
		panel.add(lblNewSprdshtSaveLoc);
		
		JLabel lblNewSprdshtFilePrompt = new JLabel("File Name:");
		lblNewSprdshtFilePrompt.setBounds(35, 47, 67, 16);
		panel.add(lblNewSprdshtFilePrompt);
		
		JButton btnNewSprdshtSaveLoc = new JButton("Browse...");
		btnNewSprdshtSaveLoc.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				openSaveDialog(newSprdshtFileChooser, lblNewSprdshtSaveLoc, true);
			}
		});
		btnNewSprdshtSaveLoc.setBounds(6, 18, 91, 25);
		panel.add(btnNewSprdshtSaveLoc);
		
		fldNewSprdshtFileName = new JTextField();
		fldNewSprdshtFileName.setColumns(10);
		fldNewSprdshtFileName.setBounds(101, 44, 171, 22);
		panel.add(fldNewSprdshtFileName);
		
		JLabel lblSpreadsheetInformation = new JLabel("Spreadsheet Information");
		lblSpreadsheetInformation.setFont(new Font("Arial", Font.BOLD | Font.ITALIC, 10));
		lblSpreadsheetInformation.setBounds(97, 13, 126, 16);
		spreadsheetPanel.add(lblSpreadsheetInformation);
	}

	/**
	 * Adds a checkbox corresponding to each of the three generated documents, as well as
	 * a button to generate documents.
	 */
	private void createChkboxAndButton() {
		chckbxGenerateCommentSheet = new JCheckBox("Generate template comment sheet");
		chckbxGenerateCommentSheet.setSelected(true);
		chckbxGenerateCommentSheet.setBounds(30, 73, 266, 25);
		contentPane.add(chckbxGenerateCommentSheet);
		
		chckbxGenerateOitScan = new JCheckBox("Generate OIT Scan Sheet");
		chckbxGenerateOitScan.addItemListener(new ItemListener() {
			@Override
			public void itemStateChanged(ItemEvent arg0) {
				changePanelDisableState(oitScanInfoPanel, chckbxGenerateOitScan.isSelected());
			}
		});
		chckbxGenerateOitScan.setSelected(true);
		chckbxGenerateOitScan.setBounds(30, 95, 196, 25);
		contentPane.add(chckbxGenerateOitScan);
		
		chckbxSpreadsheet = new JCheckBox("Add class to spreadsheet");
		chckbxSpreadsheet.addItemListener(new ItemListener() {
			@Override
			public void itemStateChanged(ItemEvent arg0) {
				Boolean genSprdsht = chckbxSpreadsheet.isSelected();
				
//				changePanelDisableState(existSprdshtPanel, genSprdsht);	
//				changePanelDisableState(newSprdshtPanel, genSprdsht);
//				spreadsheetTabbedPane.setEnabled(genSprdsht);
				
				changePanelDisableState(spreadsheetPanel, genSprdsht);
				
			}
		});
		chckbxSpreadsheet.setSelected(true);
		chckbxSpreadsheet.setBounds(30, 117, 230, 25);
		contentPane.add(chckbxSpreadsheet);
		
		JButton btnGenerateDocuments = new JButton("Generate Documents");
		btnGenerateDocuments.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				DocInfo info = retreiveDataFromFields();
				
				int sprdshtChoice;
				// Determine what has to be checked for validation for spreadsheet
				if(chckbxSpreadsheet.isSelected()){
					int index = spreadsheetTabbedPane.getSelectedIndex();
					
					// Using existing spreadsheet
					if(index == EXIST_SPRDSHT_INDEX){
						sprdshtChoice = SPREADSHEET_EXISTING;
					}
					// Creating new spreadsheet
					else{
						sprdshtChoice = SPREADSHEET_NEW;
					}
				}
				// Not using spreadsheet
				else{
					sprdshtChoice = SPREADSHEET_NONE;
				}
				
				// If data is valid in fields, generate documents that have checked boxes 
				if(isValid(info, chckbxGenerateOitScan.isSelected(), 
						chckbxGenerateCommentSheet.isSelected(), sprdshtChoice )){
					generateDocuments(info);
				}
			}
		});
		btnGenerateDocuments.setBounds(99, 792, 161, 45);
		contentPane.add(btnGenerateDocuments);
		
		// Make button default, allowing user to hit enter to generate documents
		this.getRootPane().setDefaultButton(btnGenerateDocuments);
		
		
	}
}
