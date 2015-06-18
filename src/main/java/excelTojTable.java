import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;



class ExcelExporter {

    void fillData(JTable table, File file) {

        try {

            WritableWorkbook workbook1 = Workbook.createWorkbook(file);
            WritableSheet sheet1 = workbook1.createSheet("First Sheet", 0);
            TableModel model = table.getModel();

            for (int i = 0; i < model.getColumnCount(); i++) {
                Label column = new Label(i, 0, model.getColumnName(i));
                sheet1.addCell(column);
            }
            int j = 0;
            for (int i = 0; i < model.getRowCount(); i++) {
                for (j = 0; j < model.getColumnCount(); j++) {
                    Label row = new Label(j, i + 1,
                            model.getValueAt(i, j).toString());
                    sheet1.addCell(row);
                }
            }
            workbook1.write();
            workbook1.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}

public class excelTojTable extends JFrame {
	
	static JTable table; 
    static JScrollPane scroll; 
    // header is Vector contains table Column 
    static Vector headers = new Vector(); 
   // static Vector data = new Vector();
    // Model is used to construct 
    DefaultTableModel model = null; 
    // data is Vector contains Data from Excel File 
    static Vector data = new Vector();
    static JButton jbClick; 
    static JFileChooser jChooser; 
    static int tableWidth = 0;
    static int tableHeight = 0; 
    static JTable tblContacts;
public excelTojTable()
 { 
      super("Import Excel To JTable");
      setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); 
      JPanel buttonPanel = new JPanel(); 
      buttonPanel.setBackground(Color.white); 
      jChooser = new JFileChooser(); 
      jbClick = new JButton("Select Excel File"); 
      buttonPanel.add(jbClick, BorderLayout.CENTER); 

        // Show Button Click Event 
          jbClick.addActionListener(new ActionListener() 
            { 
                     //@Override 
                     public void actionPerformed(ActionEvent arg0) 
                              { 
                                     jChooser.showOpenDialog(null); 
                                     jChooser.setDialogTitle("Select only Excel workbooks");
                                     File file = jChooser.getSelectedFile();
                                    if(file==null)
                                      {
                                          JOptionPane.showMessageDialog(
                                          null, "Please select any Excel file",
                                          "Help",
                                          JOptionPane.INFORMATION_MESSAGE); 
                                          return;
                                        }
                                    else if(!file.getName().endsWith("xls"))
                                       { 
                                             JOptionPane.showMessageDialog(
                                             null, "Please select only Excel file.", 
                                            "Error",JOptionPane.ERROR_MESSAGE); 
                                       }
                                    else 
                                      { 
                                            fillData(file);
                                            model = new DefaultTableModel(data, headers); 
                                            tableWidth = model.getColumnCount() * 150; 
                                            tableHeight = model.getRowCount() * 25; 
                                            table.setPreferredSize(new Dimension( tableWidth, tableHeight));
                                             table.setModel(model); 
                                         } 
                              } 
            });
          final JButton buttonSave = new JButton("Save");
          buttonSave.setBounds(350, 325, 100, 20);
          buttonPanel.add(buttonSave);

          // Set up Save button method
          buttonSave.addActionListener(new ActionListener(){
              //@Override
              public void actionPerformed(ActionEvent e) {
                  try{

                      BufferedWriter bfw = new BufferedWriter(new FileWriter("C:\\Users\\Steve\\Workspace\\ContactInfo\\ContactInfo.txt"));

                        for (int i = 0 ; i < tblContacts.getRowCount(); i++)
                        {

                          for(int j = 0 ; j < tblContacts.getColumnCount();j++)
                          {
                              bfw.newLine();
                              bfw.write((String)(tblContacts.getValueAt(i,j)));
                              bfw.write("\t");;
                          }
                        }
                        bfw.close();
              }catch(Exception ex) {

                  ex.printStackTrace();
              }
              }
          });
         table = new JTable(); 
         table.setAutoCreateRowSorter(true); 
         model = new DefaultTableModel(data, headers);
         table.setModel(model); 
         table.setBackground(Color.pink);
         table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF); 
         table.setEnabled(false); 
         table.setRowHeight(25); 
         table.setRowMargin(4); 
         tableWidth = model.getColumnCount() * 150; 
         tableHeight = model.getRowCount() * 25;
         table.setPreferredSize(new Dimension( tableWidth, tableHeight)); 
         scroll = new JScrollPane(table);
         scroll.setBackground(Color.pink);
         scroll.setPreferredSize(new Dimension(300, 300)); 
         scroll.setHorizontalScrollBarPolicy( JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
         scroll.setVerticalScrollBarPolicy( JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED); 
         getContentPane().add(buttonPanel, BorderLayout.NORTH); 
         getContentPane().add(scroll, BorderLayout.CENTER); 
         setSize(600, 600); 
         setResizable(true);
         setVisible(true); 
 } 

// Fill JTable with Excel file data. * * @param file * file :contains xls file to display in jTable 

  void fillData(File file) 
      { 
         int index=-1;
         HSSFWorkbook workbook = null; 
        try { 
               try { 
                       FileInputStream inputStream = new FileInputStream (file);
                        workbook = new HSSFWorkbook(inputStream);
                    } 
               catch (IOException ex) 
                    { 
                         Logger.getLogger(excelTojTable.class. getName()).log(Level.SEVERE, null, ex);
                     } 

                       String[] strs=new String[workbook.getNumberOfSheets()];
                      //get all sheet names from selected workbook
                        for (int i = 0; i < strs.length; i++) {
                             strs[i]= workbook.getSheetName(i); }
                        JFrame frame = new JFrame("Input Dialog");
                      
                        String selectedsheet = (String) JOptionPane.showInputDialog(
                           frame, "Which worksheet you want to import ?", "Select Worksheet",
                          JOptionPane.QUESTION_MESSAGE, null, strs, strs[0]);
                
                       if (selectedsheet!=null) {
                            for (int i = 0; i < strs.length; i++)
                              {
                                 if (workbook.getSheetName(i).equalsIgnoreCase(selectedsheet))
                                 index=i; }
                            HSSFSheet sheet = workbook.getSheetAt(index);
                            HSSFRow row=sheet.getRow(0);
                        
                           headers.clear();
                           
                           System.out.println("the row values::::::::;"+row.getLastCellNum());
                           
                          // System.out.println("the row values::::::::;"+row.getFirstCellNum());
                           
                           //System.out.println("the row values::::::::;"+row.getRowNum());
                           for (int i = 0; i < row.getLastCellNum(); i++)
                          {
                        	   
                        	 //  System.out.println("the row values in loop::::::::;"+row.getLastCellNum());
                             HSSFCell cell1 = row.getCell(i);
                             
                             //System.out.println("the cell values::::::::;"+row.);
                             headers.add(cell1.toString());
                             
                            // System.out.println("headers.add(cell1.toString())::::::::;"+headers.add(cell1.toString()));
                          }
                        
                          data.clear();
                          for (int j = 1; j < sheet.getLastRowNum() + 1; j++)
                          {
                             Vector d = new Vector();
                             row=sheet.getRow(j);
                             int noofrows=row.getLastCellNum();
                             System.out.println("the row noofrows in loop::::::::;"+noofrows);
                             
                             for (int i = 0; i < noofrows; i++)
                             {    //To handle empty excel cells 
                                   HSSFCell cell=row.getCell(i,org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK  );
                                   System.out.println("the row values in loop::::::::;"+row.getCell((short)i ));
                                  d.add(cell.toString());
System.out.println("value of d:::::::::"+d.firstElement());
                             }
                            d.add("\n");
                            data.add(d);
                          }
                     }
                       
                       
                    else { return; }
        
                       //JFrame frame = new JFrame("JTable to Excel");
                       DefaultTableModel model = new DefaultTableModel(data, headers);
                       final JTable table = new JTable(model);
                       JScrollPane scroll = new JScrollPane(table);

                       JButton export = new JButton("save");
                       export.addActionListener(new ActionListener() {

                          // @Override
                           public void actionPerformed(ActionEvent evt) {

                               try {
                                   ExcelExporter exp = new ExcelExporter();
                                   exp.fillData(table, new File("/home/ram/result.xls"));
                                   JOptionPane.showMessageDialog(null, "Data saved at " +
                                           "'home: \\ result.xls' successfully", "Message",
                                           JOptionPane.INFORMATION_MESSAGE);
                               } catch (Exception ex) {
                                   ex.printStackTrace();
                               }
                           }
                       });

                       frame.getContentPane().add("Center", scroll);
                       frame.getContentPane().add("South", export);
                       frame.pack();
                       frame.setVisible(true);
                       frame.setDefaultCloseOperation(frame.EXIT_ON_CLOSE);}
      catch (Exception e) { e.printStackTrace(); } }
  
public static void main(String[] args)
       { 
		excelTojTable et=new excelTojTable();
		
		
       }
   }
