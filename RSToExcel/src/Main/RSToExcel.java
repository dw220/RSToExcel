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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class RSToExcel 
{
	
	private RSToExcel()
	{		
		String error = "RSToExcel is meant to run as a static class";
		throw new RuntimeException(error);		
	}	
	
	public static void Covert( ResultSet rs, String saveTo, boolean open, String fileName ) 
	throws SQLException, IOException
	{		
		// create file name
		String outputFile = saveTo + fileName + ".xlsx";
		
		ArrayList<String> colNames = getColumnNames( rs );		
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
			// get amount of columns in the record set
			int rsColCount = rs.getMetaData().getColumnCount();			

			// only set if there are records in the set
			if( rsColCount > 1 ){ colVals = new ArrayList<String>(); }
			else{ return null; }
			
			// start getting all the column values
			for( int i=1; i<rsColCount; i++){
				String currCol  = rs.getMetaData().getColumnName(i);
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
	 * Optional 
	 * @param workbook
	 * @return
	 */
	@SuppressWarnings("unused")
	private static XSSFWorkbook createFileHeaders( XSSFWorkbook workbook )
	{		
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
		CellStyle style = workbook.createCellStyle();
		
		// set the font
		style.setFont( font );
		
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
		// get the metadata associated with the recordset
		ResultSetMetaData rsMeta = rs.getMetaData();
		
		// get fcolumn count
		int cols = rsMeta.getColumnCount();
		
		ArrayList<String> colTypes = new ArrayList<String>();
		
		// add the types from the meta data
		for( int i=1; i<cols; i++ ){    
			colTypes.add( rsMeta.getColumnClassName(i) );
		}
		
		return colTypes;		
	}	
	
	private static void addData( ResultSet rs, XSSFWorkbook workbook ) 
			throws SQLException
	{	
		// get the meta data
		ResultSetMetaData rsMeta = rs.getMetaData();
		int cols = rsMeta.getColumnCount();
		
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
			for( int i=1; i<cols; i++ )
			{
				// get the object so we can check for null
				Object obj = rs.getObject( i );
				if (!( obj == null)) {							
					curVal = obj.toString();
				}
				
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
		String cmdString = "rundll32 url.dll,FileProtocolHandler ";
		
		if( new File(filePath).exists() )
		{
			try{
			Process p = Runtime
					.getRuntime()
					.exec( cmdString + filePath);
			p.waitFor();
			} catch( Exception e ){
				e.printStackTrace();
				return false;
			}
		}
		
		return true;
		
	}
	
	
	
	
}
