import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;


public class ExcelReader extends Exception{
	
	public static Recordset recordset;
	public static Fillo fillo = new Fillo();
	public static File fl = new File("C:\\Users\\soham\\Automation_Testing\\ReadExcel\\src\\ColumnListfromexcel.properties");
	

	public static void main(String[] args) throws Exception {

		ArrayList<String> fieldnames = listtheColumnnames();
		System.out.println("field from main : " + fieldvalues(15));	
	}
	
	// get row data as per the row
	public static Map<String, String> fieldvalues(int rowcount) throws Exception {
		String query = "select * from " + propertiesExternal("excelSheetname");
		Connection connection = getFilloconnection();
		Recordset rs = connection.executeQuery(query);
		Map<String, String> rowmap = new HashMap<String, String>();
		int counter = 1;
		// As the number of the properties will always be same as the no of columns
		int noofproperties = noOfcolumns();
		System.out.println("no of props : " + noofproperties);
		while (rs.next()) {
			if (counter == rowcount) {
				for (int i = 0; i < noofproperties; i++) {
					// prepare the key
					String prepareKey = "keys"+i;
					rowmap.put(prepareKey, rs.getField(properties("key"+i)));
				}
			}
			counter++;
		}
		return rowmap;
	}
	
	// delete the property file
	public static void deleteproperty() {
	    if(fl.delete()) 
            { 
            System.out.println("File deleted successfully"); 
            } 
            else
            { 
            System.out.println("Failed to delete the file"); 
            }
	}
	
	// Read The property file
	public static String properties(String key) throws Exception {
		FileInputStream file = new FileInputStream(fl);
		Properties rpop = new Properties();
		rpop.load(file);
		String data = rpop.getProperty(key);
		return data;
	}
	
	// Read The property file FROM EXTERNAL USER
		public static String propertiesExternal(String key) throws Exception {
			File files = new File("C:\\Users\\soham\\Automation_Testing\\ReadExcel\\Config\\config.properties");
			FileInputStream file = new FileInputStream(files);
			Properties rpop = new Properties();
			rpop.load(file);
			String data = rpop.getProperty(key);
			return data;
		}
	
	// Counts the column
	public static int noOfcolumns() throws Exception {
		String query = "select * from " + propertiesExternal("excelSheetname");
		Connection connection = getFilloconnection();
		Recordset rs = connection.executeQuery(query);
		int count = rs.getFieldNames().size();
		return count;
	}
	
	// lists the column names and writes it it external properties file 
	public static ArrayList<String> listtheColumnnames() throws Exception {
		String query = "select * from " + propertiesExternal("excelSheetname");
		Connection connection = getFilloconnection();
		Recordset rs = connection.executeQuery(query);
		ArrayList<String> elements = rs.getFieldNames();
		Properties properties = new Properties();
		for (int i = 0; i < elements.size(); i++) {
			String key = "key" + i;
			properties.setProperty(key, elements.get(i));
		}
		FileOutputStream fileOutputStream = new FileOutputStream(fl);
		properties.store(fileOutputStream, "Columns from the excel sheet");
		fileOutputStream.close();
		return elements;
	}
	
	// Counts the number of rows
	public static int noOfrows() throws Exception{
		String query = "select * from " + propertiesExternal("excelSheetname");
		Connection connection = getFilloconnection();
		int rowcount = connection.executeQuery(query).getCount();
		return rowcount;
	}
	
	// Gets the fillo connection
	public static Connection getFilloconnection() throws Exception {
		Connection connection = fillo.getConnection(propertiesExternal("excelFilepath"));
		return connection;
	}

}
