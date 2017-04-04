package Test;

import static org.junit.Assert.*;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.junit.Before;

import Main.RSToExcel;

public class Test {
	
	Connection conn;
	
	@Before
	public void initConn() 
			throws ClassNotFoundException, SQLException 
	{
		// Load class
		Class.forName("org.postgresql.Driver");
		
		// Get a new connection
		conn = DriverManager.getConnection( "jdbc:postgresql://tp-dev-01:5432/saracs_development" , "postgres" , "P)9oI*7uY^" );
	}

	@org.junit.Test
	public void test()  
	{		
		String SQL = "select * from \"Printers\";";
		
		try
		{
			// create the rs to convert to excel spreadsheet
			ResultSet rs = conn.createStatement().executeQuery(SQL);
			
			// convert
			RSToExcel.Covert( rs , "e:\\DEV07\\", true, "output");
			
		} catch(Exception e)
		{
			e.printStackTrace();
			fail("Connection / rs loading");			
		}
	}
	


}
