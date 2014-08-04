/** 
 * Simple class used to store the data for a    
 * class and faculty support staff. It assumes the data   
 * has already been validated.							  
 */

package org.EvalGenerator;

public class DocInfo {
	//******************* DATA MEMBERS *******************
	private String instFName;
	private String instLName;
	private String subject;
	private String courseNum;
	private String section;
	private String year;
	private String facSuppName;
	private String facSuppExten;
	private String mailbox;
	private Semester semester;
	
	public enum Semester { Fall, Winter, Spring, Summer };
	
	//******************* CONSTRUCTORS *******************
	/**
	 * Parameterized constructor for DocInfo. Assumes all values 
	 * have been checked for type errors and are not null.
	 * @param fname First name of instructor
	 * @param lname Last name of instructor
	 * @param sub Class subject (ex. CST)
	 * @param cNum Class number
	 * @param sec Class section (ex. WN123)
	 * @param yr Class year
	 * @param fsName Faculty support name (First and last)
	 */
	public DocInfo(String fname, String lname, String sub,
			String cNum, String sec, String yr, String fsName, 
			String fsExten, String mail, Semester sem){
		
		instFName    = fname;
		instLName    = lname;
		subject      = sub;    
		courseNum    = cNum;  
		section      = sec;    
		year         = yr;
		facSuppName  = fsName;		
		facSuppExten = fsExten;
		mailbox      = mail;
		semester     = sem;
	}
	
	//******************* PUBLIC METHODS *******************
	public String getInstFName(){
		return instFName;
	}
	public String getInstLName(){
		return instLName;
	} 
	public String getSubject(){
		return subject;
	}
	public String getCourseNum(){
		return courseNum;
	}
	public String getSection(){
		return section;
	}
	public String getYear(){
		return year;
	}
	public String getFacSuppName(){
		return facSuppName;
	}
	public String getFacSuppExten(){
		return facSuppExten;
	}
	public String getMailbox(){
		return mailbox;
	}
	public Semester getSemester(){
		return semester;
	}
	
	public void setInstFName(String instFName) {
		this.instFName = instFName;
	}

	public void setInstLName(String instLName) {
		this.instLName = instLName;
	}

	public void setSubject(String subject) {
		this.subject = subject;
	}

	public void setCourseNum(String courseNum) {
		this.courseNum = courseNum;
	}

	public void setSection(String section) {
		this.section = section;
	}

	public void setYear(String year) {
		this.year = year;
	}

	public void setFacSuppName(String facSuppName) {
		this.facSuppName = facSuppName;
	}

	public void setFacSuppExten(String facSuppExten) {
		this.facSuppExten = facSuppExten;
	}

	public void setMailbox(String mailbox) {
		this.mailbox = mailbox;
	}

	public void setSemester(Semester semester) {
		this.semester = semester;
	}

	public String toString(){
		return "FirstName: " + instFName + "\n" + 
				"instLName: " + instLName + "\n" + 
				"subject: " + subject + "\n" + 
				"courseNum: " + courseNum + "\n" + 
				"section: " + section + "\n" + 
				"year: " + year + "\n" + 
				"facSuppName: " + facSuppName + "\n" + 
				"semester: " + semester + "\n";
	}
	
	//******************* PRIVATE METHODS *******************
}
