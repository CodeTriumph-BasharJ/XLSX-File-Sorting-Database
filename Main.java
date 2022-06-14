/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *Program that accepts XLSX file and read cell by cell the info provided 
 *then re-arrange them in an XML file after choosing to sort the data
 *by first name, gender, salary, or working years.
 * 
 * @author Bashar Jirjees
 */
import java.util.ArrayList;
import java.util.Scanner;
import java.io.File;
import java.io.FileWriter;
import java.io.FileNotFoundException;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.Iterator;
import java.util.Comparator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    
    static ArrayList<Cell> labels = new <Cell> ArrayList();
    static ArrayList<Cell> data = new <Cell> ArrayList();
    static int labels_counter = 0;
    static int counter = 0;
   
    /**
     * main method to identify XLSX file and read it.
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
        // TODO code application logic here
        String file_name = "MOCK_DATA.xlsx";
        File file = new File(file_name);
        if (file.exists()) {
            ReadXLSX(file_name);
            make_Choice();
        } else {
            
            System.out.println("FILE NOT FOUND!!");
            System.exit(0);
        }

    }
    
    
   /**
    * ReadXLSX method to read the XLSX file and count the number of columns 
    * available.
    */
    public static void ReadXLSX(String file_name) throws FileNotFoundException, IOException{

        FileInputStream file = new FileInputStream(new File (file_name));
        Workbook xlsxFile = new XSSFWorkbook(file);
        Sheet sheet = xlsxFile.getSheetAt(0);
        Iterator<Row> iter_1 = sheet.iterator();

        while (iter_1.hasNext()) {

            Row row = iter_1.next();
            Iterator<Cell> iter_2 = row.cellIterator();

            while (iter_2.hasNext()) {
                Cell cell = iter_2.next();

                if (counter == 0) {
                    ++labels_counter;
                    labels.add(cell);
                } else {

                    data.add(cell);
                }
            }
            ++counter;

        }

    }

    
    /**
    * XMLdata method to create the XML file and write the specified data in it.
    */
    public static void XMLdata() throws IOException {

        int index = 0;
        counter = 0;
        
        FileWriter filewriter = new FileWriter("report.xml");
        filewriter.write("<Full_Data>\n\n");
        filewriter.write("<Employee>\n");
        for (Cell c : data) {

            filewriter.write("    <" + labels.get(counter).toString().replace(" ", "_").replace("(","_").replace(")","") + ">");
            filewriter.write(c.toString());
            filewriter.write("</" + labels.get(counter).toString().replace(" ", "_").replace("(","_").replace(")","") + ">\n");
            ++index;
            ++counter;
            if (counter == labels_counter) {
                filewriter.write("</Employee>\n\n");
                if (index < data.size()) {
                    filewriter.write("<Employee>\n");
                }
                counter = 0;
            }
        }
        filewriter.write("</Full_Data>");
        filewriter.close();
    }
    
    
    /**
    * make_choice method to read the user input choice regarding the type of
    * data sorting or non-sorting.
    */
    public static void make_Choice()throws IOException{
        
        System.out.println("How would you like to sort data:");
        System.out.println("Press (1) to sort first names alphabatically.");
        System.out.println("Press (2) to sort by salary.");
        System.out.println("Press (3) to sort by working years.");
        System.out.println("Press (4) to sort by gender.");
        System.out.println("Press (0) to not sort and exit.");
        
        Scanner number_input = new Scanner(System.in);
        System.out.print("Enter your choice: ");
        int choice = number_input.nextInt();
        
        if(choice == 1) sort_Data_Alphabatically();
        if(choice == 2) sort_Data_By_Salary();
        if(choice == 3) sort_Data_By_Working_Years();
        if(choice == 4) sort_Data_By_Gender();
        if(choice == 0)  XMLdata();
 
    }
    
    
    /**
    * sort_Data_Alphabatically method to sort the employees/persons by first
    * names upon the user request.
    */
    public static void sort_Data_Alphabatically() throws IOException{
        counter = -1;
        int counter_2 = 0;
        int position = 0;
        
        for(Cell cell: labels){
            ++counter;
            ArrayList <String> first_names = new <String> ArrayList();
            
            if (cell.toString().toLowerCase().replace("_", " ").equals("first name")){
            
                while(counter < data.size()){
                if(counter_2 == 0) position = counter;
                ++counter_2;
                first_names.add(data.toArray()[counter].toString());
                counter+=labels_counter;
                
                }
            sort_store_Data(first_names, position);
            XMLdata();
            
        }else if(counter == labels_counter){
                System.out.println("Can't sort alphabatically. First names don't exist.");
                make_Choice();
                break;
            }
        }   
    }
    
    
    /**
    * sort_Data_By_Salary method to sort the employees/persons by salary in
    * increasing order upon the user request.
    */
    public static void sort_Data_By_Salary() throws IOException{

       counter = 0;
       int position = 0;
       ArrayList <String> salaries = new <String> ArrayList();
       while(counter < labels_counter)if(labels.get(counter++).toString().toLowerCase().equals("salary"))break;
       if(counter == labels_counter){
           System.out.println("Salary data isn't available.");
           make_Choice();
       }
       --counter;
       position = counter;
       while(counter < data.size()){
           salaries.add(data.get(counter).toString());
           counter+=labels_counter;
       }
      sort_store_Data(salaries, position);
       
       XMLdata();
    }
    
    
    /**
    * sort_Data_By_Gender method to sort the employees/persons by gender
    * upon the user request. It is worth noting that it sorts by putting similar 
    * genders together and starting from the first gender read in the XLSX file.
    */
    public static void sort_Data_By_Gender()throws IOException{
        
       counter = 0;
       int position = 0;
       ArrayList <String> genders = new <String> ArrayList();
       while(counter < labels_counter) if(labels.get(counter++).toString().toLowerCase().contains("gender"))break;
       if(counter == labels_counter){
           System.out.println("working years data isn't available.");
           make_Choice();
       }
       --counter;
       position = counter;
       while(counter < data.size()){
           genders.add(data.get(counter).toString());
           
           counter+=labels_counter;
           
       }
      
      sort_store_Data(genders, position);   
        
      XMLdata();
    }
    
    
    /**
    * sort_Data_By_Working_Years method to sort the employees/persons 
    * increasingly by employment period length in years upon the user
    * request.
    */
    public static void sort_Data_By_Working_Years()throws IOException{
        
       counter = 0;
       int position = 0;
       ArrayList <String> working_period = new <String> ArrayList();
       while(counter < labels_counter) if(labels.get(counter++).toString().toLowerCase().contains("years"))break;
       if(counter == labels_counter){
           System.out.println("working years data isn't available.");
           make_Choice();
       }
       --counter;
       position = counter;
       while(counter < data.size()){
           working_period.add(data.get(counter).toString());
           
           counter+=labels_counter;
           
       }
  
      sort_store_Data(working_period, position);
        
      XMLdata();
    }
    
    
    /**
     * sort_store_Data method to check which parameter has been chosen for 
     * sorting by the user and finding all the data associated with that
     * parameter then sorting them accordingly.
    */
   public static void sort_store_Data(ArrayList <String> info, int position)throws IOException{
       counter = 0;
       final int temp = position;
       int counter_2 = 0;
       int counter_3 = 0;
       int counter_4 = 0;   
       
       if(info.get(0).toString().matches(".*[a-z].*")) Collections.sort(info);
       else sort_Number(info);
       
       ArrayList <Cell> arr = new <Cell> ArrayList();
  
       while(!data.isEmpty()){
           
       while(counter < data.size() && counter_2 < info.size() && data.get(counter)!= null && (data.get(counter + temp).toString().equals(info.get(counter_2)))){
            
        
            position = counter;
            
            while (counter_3 < labels_counter){
                arr.add(data.get(position));
                data.set(position, null);
                ++position;
                ++counter_3; 
                
            }
            
            counter  = counter_3 = position = 0;
            ++counter_2;
            ++counter_4;
            break;
       }
       
       if(counter_4 == 0) counter+=labels_counter;
       if(counter_2 == info.size())break;
       counter_4 = 0;
      } 
      data = arr; 
   } 
   
   
   /**
   * sort_Number helper method used with the help of the built-in library 
   * Comparator class to compare numerical data only.
   */
   public static void sort_Number(ArrayList<String> info){
       
    Collections.sort(info, new Comparator<String>() {
        
        @Override
        public int compare(String temp1, String temp2) {
            
            return Double.valueOf(temp1).compareTo(Double.valueOf(temp2));
            
        }
    });
   }
   
   
 }