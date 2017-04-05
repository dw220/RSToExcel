import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.channels.FileChannel;
import java.nio.file.Paths;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import javax.imageio.ImageIO;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class C 
{	
	static private int colCount = 0;
	static private ResultSetMetaData rsMeta = null;
	
	private C()
	{
		String err = "Object creation is not allowed on this class";
		throw new RuntimeException( err ); 
	}
	
	/**
	 * Method for trying and parsing various date records
	 * this should be changed to be more efficent as having
	 * multiple try catch bocks are quite expensive
	 * 
	 * It works by systematiccaly trying to parse a date string
	 * using different formats until the correct one is found
	 * 
	 * @param date
	 * @return
	 */
	public static LocalDate tryParse( String date )
	{		
		LocalDate ld = null;
		
		try	{  ld = LocalDate.parse( date , DateTimeFormatter.ofPattern( "dd-MM-yyyy" ) );return ld;} 
		catch( Exception e ) {};
		
		try	{  ld = LocalDate.parse( date , DateTimeFormatter.ofPattern( "dd-MMM-yyyy" ) );return ld;} 
		catch( Exception e ) {};
		
		try	{  ld = LocalDate.parse( date , DateTimeFormatter.ofPattern( "MMM-dd-yyyy" ) );return ld;} 
		catch( Exception e ) {};
		
		try	{  ld = LocalDate.parse( date , DateTimeFormatter.ofPattern( "MM-dd-yyyy" ) );return ld;} 
		catch( Exception e ) {};	
		
		try	{  ld = LocalDate.parse( date , DateTimeFormatter.ofPattern( "yyyy-MMM-dd" ) );return ld;} 
		catch( Exception e ) {};
		
		try	{  ld = LocalDate.parse( date , DateTimeFormatter.ofPattern( "yyyy-MM-dd" ) ) ;return ld;} 
		catch( Exception e ) {};
				
		return ld;		
	}
	
	/**
	 * Returns the current date
	 * @return
	 */
	public static LocalDate currentDate() {		
		return LocalDate.now();		
	}
	
	/**
	 * Indents quotes to a string, example:
	 * 
	 * String s = C.iQuote( "Hello World" )
	 * 
	 * s will be equal to "Hello World" // Quotes are apart of string
	 * 
	 */
	public static String addQuotes(String s){
		return "\" "+ s + "\"";		
	}
	
	/**
	 * 
	 * Creates a timestamp of the current time and date
	 * @deprecated This should not be used, timestamps should be created in the DB
	 * @return
	 * 
	 */
	public static Timestamp getTimeStamp()
	{
		// Calendar information
		Calendar calendar 		= Calendar.getInstance();
		java.util.Date now 		= calendar.getTime();
		Timestamp dbStamp 		= new Timestamp(now.getTime());
		return dbStamp;
	}	
	
	/**
	 * works with format hh:mm:ss
	 * @param pTime
	 * @return
	 */
	public static String StringToTime( String pTime )
	{		
		if( pTime.length() < 4  ){
			return "";
		}		
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern( "hh:mm:ss" );
		return LocalTime.parse(pTime, dtf ).toString();
	}	

	/**
	 * 
	 * @param s
	 * @return
	 */
	public static String emptyStringToNull( String s ){
		if( s.length() == 0 )
			return null;
		else
			return s;
	}	
	
	/**
	 * This method checks to see if we are in the debugger or not
	 * please remeber you being in the IDE is not the same as being
	 * in the IDE and executing the program by pressing the DEBUG button
	 * @return
	 */
	public static Boolean IDE()
	{
		boolean isDebug = java.lang.management.ManagementFactory.getRuntimeMXBean().
		getInputArguments().toString().indexOf("-agentlib:jdwp") > 0;
		return isDebug;
	}
	
	/**
	 * 
	 * @param rs
	 * @param saveTo
	 * @param open
	 * @param fileName
	 * @throws SQLException
	 * @throws IOException
	 */
	public static void Convert( ResultSet rs, String saveTo, boolean open, String fileName ) 
	throws SQLException, IOException
	{		
		// set Class variables
		colCount = rs.getMetaData().getColumnCount();
		rsMeta = rs.getMetaData();
		
		
		// create file name
		String outputFile = saveTo + fileName + ".xlsx";
		
		// get all column names from the record set
		ArrayList<String> colNames = getColumnNames( rs );
		
		// create file
		XSSFWorkbook workbook = createFile( );
		
		// DW create cell header? Including time/date
		//workbook = createFileHeaders(workbook);		
		
		// set the column headers from a recordset
		workbook = createColHeaders( workbook, colNames );
		
		// get column data types - this currently isn't needed
		// but here for future development
		getColDataTypes( rs );
		
		// add the recordset data
		addData( rs, workbook );
				
		// save the workbook to a certain file
		saveSheet(outputFile, workbook);
		
		// open file after creation and saving
		if ( open )
			openFile( outputFile );
	}
	
	private static ArrayList<String> getColumnNames( ResultSet rs )
	{		
		// list to store all column names
		ArrayList<String> colVals = null;		
		try 
		{	
			// only set if there are records in the set
			if( colCount > 1 ){ colVals = new ArrayList<String>(); }
			else{ return null; }
			
			// start getting all the column values
			for( int i=1; i<colCount; i++){
				String currCol  = rsMeta.getColumnName(i);
				colVals.add( currCol );
			}
			return colVals;
			
		} catch (SQLException e) {	
			e.printStackTrace(); 
		}

		return null;
	}
	
	/**
	 * Create the work book we are going to use
	 * @param saveTo
	 * @return
	 * @throws IOException 
	 */
	private static XSSFWorkbook createFile(  ) throws IOException
	{
		//DW create a new workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		//DW create a sheet
		XSSFSheet sheet = workbook.createSheet();
		return workbook;
	}
	
	/**
	 * 
	 * Save the workbook to the specified location
	 * 
	 * @param outputFile
	 * @param workbook
	 * @return
	 * @throws IOException
	 */
	private  static boolean saveSheet(String outputFile, XSSFWorkbook workbook) 
			throws IOException
	{
		
		// convert output String to a file
		File f = Paths.get( outputFile ).toFile();
		
		// check if it exists, if so then we need to delete it
		if( f.exists() ){ f.delete(); }
		
		// save to the disk
		FileOutputStream os = new FileOutputStream(outputFile);

		// write work book to file
		workbook.write( os );
		return true;
	}
	
	/**
	 * @param workbook
	 * @return
	 */
	@SuppressWarnings("unused")
	private static XSSFWorkbook createFileHeaders( XSSFWorkbook workbook )
	{	
		// Create Header
		String header = "Created:" + LocalDateTime.now().toString();
		
		XSSFSheet sheet = workbook.getSheetAt( 0 );
		XSSFRow row = sheet.createRow(0);
		
		Cell c = row.createCell( 0 );
		
		c.setCellValue( header  );
		
		return workbook;
	}
	
	/**
	 * 
	 * Add all the column headers from the result set to the 
	 * Excel spreadsheet
	 * 
	 * @param workbook
	 * @param colNames
	 * @return
	 */
	private static XSSFWorkbook createColHeaders(XSSFWorkbook workbook, ArrayList<String> colNames)
	{	
		// get the sheet
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		// create a row to store the column headers
		XSSFRow row = sheet.createRow( 0 );
		
		// Create header style for cells
		Font font = workbook.createFont();
		font.setFontHeightInPoints( (short)16 );
		font.setBoldweight( (short)16 );
		
		// create the cell style
		XSSFCellStyle  style = workbook.createCellStyle();
		
		// set the font
		style.setFont( font );
		
		// set background color, and type of fill
		style.setFillForegroundColor( new XSSFColor(new java.awt.Color(193, 217, 255)) );
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		// set border on header columns
		style.setBorderLeft( BorderStyle.MEDIUM );
		style.setBorderRight( BorderStyle.MEDIUM );
		style.setBorderTop( BorderStyle.MEDIUM );
		style.setBorderBottom( BorderStyle.MEDIUM );
				
		// loop through and write the columns
		for( int i=0; i < colNames.size(); i++ )
		{
			Cell c = row.createCell( i );
			c.setCellValue( colNames.get( i ) );
			c.setCellStyle( style );
			
			// auto size column based on header
			// this is no inefficent to do with every cell
			sheet.autoSizeColumn( i );
		}
		
		return workbook;
	}
	
	/**
	 * 
	 * Get all column data types, for example VARCHAR 
	 * will return Java.Lang.String, int will return integer etc...
	 * 
	 * @param rs
	 * @return
	 * @throws SQLException
	 */
	public static ArrayList<String> getColDataTypes( ResultSet rs ) 
			throws SQLException
	{		
		ArrayList<String> colTypes = new ArrayList<String>();
		
		// add the types from the meta data
		for( int i=1; i<colCount; i++ ){    
			colTypes.add( rsMeta.getColumnClassName(i) );
		}
		
		return colTypes;		
	}	
	
	private static void addData( ResultSet rs, XSSFWorkbook workbook ) 
			throws SQLException
	{	
		// count is needed as it refers to the column in the RS
		int count = 0;
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		// if SQL null is returtned this will be used on the sheet
		String curVal = "SQL NULL";
		while( rs.next() )
		{			
			count++;
			
			// create a new row references a row in RS using count
			Row r = sheet.createRow( count );
			for( int i=1; i<colCount; i++ )
			{
				// get the object so we can check for null
				Object obj = rs.getString( i );
				if (!( obj == null)) {							
					curVal = obj.toString();
				} else
					curVal = "NULL";
				
				// cell positions start at 0, rows start at 1
				// so this is needed
				Cell c = r.createCell( i-1 );
				
				// set the cell value
				c.setCellValue( curVal );
			}
		}		
		
	}
	
	/**
	 * Shells out and open the file with the run time
	 * @param filePath
	 * @return
	 */
	public static boolean openFile( String filePath ) 
	{		
		// DW cmdString to shell out with, this opens a file
		String cmdString = "rundll32 url.dll,FileProtocolHandler " + filePath;	
		if( new File(filePath).exists() )
		{
			try
			{
				Process p = Runtime.getRuntime().exec( cmdString);
				p.waitFor();
			} catch( Exception e )
			{
				e.printStackTrace();
				return false;
			}
		}		
		return true;		
	}
	
	/**
	 * Fast File copy mechanism - moves one file to another
	 * @param in
	 * @param out
	 * @throws IOException
	 */
	public static void fileCopy( File in, File out )
            throws IOException
    {
        FileChannel inChannel = new FileInputStream( in ).getChannel();
        FileChannel outChannel = new FileOutputStream( out ).getChannel();
        try
        {
            // magic number for Windows, 64Mb - 32Kb)
            int maxCount = (64 * 1024 * 1024) - (32 * 1024);
            long size = inChannel.size();
            long position = 0;
            while ( position < size )
            {
               position += inChannel.transferTo( position, maxCount, outChannel );
            }
        }
        finally {
            if ( inChannel != null ) { inChannel.close(); }
            if ( outChannel != null ){ outChannel.close();}
        }
    }
	
	/**
	 * 
	 * Caputres a screen shot for use with debugging
	 * 
	 * @param fileName
	 * @throws Exception
	 */
	public void captureScreen(String fileName) 
			throws Exception 
	{
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		Rectangle screenRectangle = new Rectangle(screenSize);
		Robot robot = new Robot();
		BufferedImage image = robot.createScreenCapture(screenRectangle);
		ImageIO.write(image, "png", new File(fileName));
	}
	
	
}