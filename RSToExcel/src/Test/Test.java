package Test;

import static org.junit.Assert.*;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;

import javax.swing.JOptionPane;

import org.junit.Before;

import Main.RSToExcel;

public class Test {
	
	Connection conn;
	
	@Before
	public void initConn() 
			throws ClassNotFoundException, SQLException 
	{

	}

	@org.junit.Test
	public void test()  
	{		
		long start = System.nanoTime();
		String SQL = "select * from \"t Users\";";
		
		try
		{
			// create the rs to convert to excel spreadsheet
			ResultSet rs = conn.createStatement().executeQuery(SQL);
			// convert
			RSToExcel.Convert( rs , "e:\\DEV07\\", true, "output");
			
		} catch(Exception e)
		{
			e.printStackTrace();
			fail("Connection / rs loading");			
		}
		long estimatedTime = System.nanoTime() - start;
		
		System.out.println( "Time taken: " + estimatedTime  / 1000000000.0 );
		
	}


}
