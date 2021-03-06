package Main;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.time.LocalDateTime;
import java.util.ArrayList;

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

/**
 * Created by DW / Development finished on 04/04/2017
 * 
 * TODO: 
 * 
 * 1) Add support for closing a file if it is currently being used
 * 2) Add support for custom styling of cells
 * 3) Set the correct data types in columns currently this outputs to string
 * 4) Get result set metadata can be set globaly
 * 
 * @author DarrenW
 */

public final class RSToExcel 
{
	
	static private int colCount = 0;
	static private ResultSetMetaData rsMeta = null;
	/**
	 * Set to private to ensure that an instance can not be create
	 */
	private RSToExcel()
	{		
		String error = "RSToExcel is meant to run as a static class";
		throw new RuntimeException(error);		
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
	 * Shell out and open the file with the run time
	 * @param filePath
	 * @return
	 */
	private  static boolean openFile( String filePath ) 
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
	
	
	
	
}
