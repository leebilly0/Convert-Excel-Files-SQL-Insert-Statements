import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.util.Iterator;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Convert .XLSX Excel file to .SQL
 * Has a GUI to open and select file
 * SQL file will be saved in the same path
 * @author Billy
 *
 */
public class Program1 extends JPanel implements ActionListener 
{
	private JFileChooser fileChooser;
	private JButton	open;
	private File file1;
	private BufferedReader br;
	int returnValue;
	String currentLine;

	public Program1 () {
		fileChooser = new JFileChooser(System.getProperty("user.dir"));
		open = new JButton("Open");

		setPreferredSize(new Dimension(250, 150));
		setLayout(null);

		add(open);

		open.setBounds(75, 100, 100, 25);
		open.addActionListener(this);

	}

	public void actionPerformed(ActionEvent e) {
		if(e.getSource() == open) 
		{
			returnValue = fileChooser.showOpenDialog(null);
			if(returnValue == JFileChooser.APPROVE_OPTION) 
			{
				file1 = fileChooser.getSelectedFile();



				//Read file and print to console
				try
				{
					FileInputStream file = new FileInputStream(new File("input.xlsx"));
					FileOutputStream fos = new FileOutputStream("output.sql");
					boolean firstTime = true;
					String outputString = "INSERT INTO 'tbl_events' ('id', ";
					int rowCount = 1;

					//Create Workbook instance holding reference to .xlsx file
					XSSFWorkbook workbook = new XSSFWorkbook(file);

					//Get first/desired sheet from the workbook
					XSSFSheet sheet = workbook.getSheet("assignment 1 sample data");

					//Iterate through each rows one by one
					Iterator<Row> rowIterator = sheet.iterator();

					

					//iterates through rows
					while (rowIterator.hasNext())
					{
						Row row = rowIterator.next();
						//For each row, iterate through all the columns
						Iterator<Cell> cellIterator = row.cellIterator();
						
						//if this is more rows being added to the table, put front parenthesis
						if (firstTime == false)
						{
							outputString = outputString + "(" + rowCount + ", ";
							rowCount++;
						}

						//iterates through columns
						while (cellIterator.hasNext())
						{
							Cell cell = cellIterator.next();
							//Check the cell type and format accordingly
							//add cell values to string
							switch (cell.getCellType())
							{
							case Cell.CELL_TYPE_STRING:
								outputString = outputString + "'" + cell.getStringCellValue() + "'";
								break;
							case Cell.CELL_TYPE_NUMERIC:
								outputString = outputString + "'" + cell.getNumericCellValue() + "'";
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								outputString = outputString + "'" + cell.getBooleanCellValue() + "'";
								break;
							case Cell.CELL_TYPE_BLANK:
								outputString = outputString + "'" + "'";
								break;
							}

							//checks to see if column for row has ended, if so put end parenthesis
							if (cellIterator.hasNext())
								outputString = outputString + ", ";
							else
								outputString = outputString + ")";
						}

						//if there are no more rows end statement
						if (!rowIterator.hasNext())
							outputString = outputString + ";";
						else
						{
							//checks to see if its finishing insert query state, if not, then put comma
							//to indicate adding more values to table
							if (firstTime)
							{
								outputString = outputString + " VALUES";
								firstTime = false;
							}
							else
								outputString = outputString + ",";

							
						}
					}
					file.close();
					fos.write(outputString.getBytes());
					fos.close();
				}
				catch (Exception e1)
				{
					e1.printStackTrace();
				}
			}
		}
	}

	public static void main(String[] args)
	{
		//Create and set up the window.
		JFrame frame = new JFrame("COPY TEXT FILE");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		//Add content to the window.
		frame.add(new Program1());
		//Display the window.
		frame.pack();
		frame.setVisible(true);
	}

}
