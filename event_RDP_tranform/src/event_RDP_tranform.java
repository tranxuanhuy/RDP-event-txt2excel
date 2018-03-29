import java.awt.EventQueue;




import javax.swing.JFrame;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import javax.swing.JButton;
import java.awt.FlowLayout;
import javax.swing.JFileChooser;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.awt.event.ActionEvent;

public class event_RDP_tranform {

	private JFrame frmRdpEventTxtexcel;
	private Workbook workbook;
	private BufferedReader br;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					event_RDP_tranform window = new event_RDP_tranform();
					window.frmRdpEventTxtexcel.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public event_RDP_tranform() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmRdpEventTxtexcel = new JFrame();
		frmRdpEventTxtexcel.setTitle("RDP-event txt2excel");
		frmRdpEventTxtexcel.setBounds(100, 100, 283, 101);
		frmRdpEventTxtexcel.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmRdpEventTxtexcel.getContentPane().setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));
		
		JButton btnOpen = new JButton("select event file (1 txt -> 1 xls)");
		btnOpen.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				 JFileChooser fileChooser = new JFileChooser();
				 fileChooser.setMultiSelectionEnabled(true);
				    int returnValue = fileChooser.showOpenDialog(null);
				    if (returnValue == JFileChooser.APPROVE_OPTION) 
				    {
				    File[] selectedFile = fileChooser.getSelectedFiles();
				    System.out.println(selectedFile);
				    
				    try {
						for (int i = 0; i < selectedFile.length; i++) {
//							dataTransform(selectedFile[i].getParent(), selectedFile[i].getName());
							String excelFilePath = selectedFile[i].getName()+".xls";
							writeExcel(selectedFile[i].getAbsolutePath(), selectedFile[i].getParent()+"\\"+excelFilePath);
						}
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				    }
			}
		});
		
		JButton btnSelectEventFile = new JButton("select event file (many txt -> 1 xls)");
		btnSelectEventFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				 JFileChooser fileChooser = new JFileChooser();
				 fileChooser.setMultiSelectionEnabled(true);
				    int returnValue = fileChooser.showOpenDialog(null);
				    if (returnValue == JFileChooser.APPROVE_OPTION) 
				    {
				    File[] selectedFile = fileChooser.getSelectedFiles();
				    System.out.println(selectedFile);
				    
				    File allContentFile = new File("NewFile.txt");
					File[] fileLocations = new File[selectedFile.length];
							
				    for (int i = 0; i < selectedFile.length; i++) {
				    	fileLocations[i]=new File(selectedFile[i].getAbsolutePath());
					}
				    mergeFiles(fileLocations, allContentFile);
				    String excelFilePath = "FormattedJavaBooks.xls";
				    try {
						writeExcel(allContentFile.getAbsolutePath(), selectedFile[0].getParent()+"\\"+excelFilePath);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				    }
			}

			private void mergeFiles(File[] files, File mergedFile) {
 
		FileWriter fstream = null;
		BufferedWriter out = null;
		try {
			fstream = new FileWriter(mergedFile, true);
			 out = new BufferedWriter(fstream);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
 
		for (File f : files) {
			System.out.println("merging: " + f.getName());
			FileInputStream fis;
			try {
				fis = new FileInputStream(f);
				BufferedReader in = new BufferedReader(new InputStreamReader(fis));
 
				String aLine;
				while ((aLine = in.readLine()) != null) {
					out.write(aLine);
					out.newLine();
				}
 
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
 
		try {
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
 
	}
		});
		frmRdpEventTxtexcel.getContentPane().add(btnSelectEventFile);
		frmRdpEventTxtexcel.getContentPane().add(btnOpen);
	}


	protected void writeExcel(String sourceFile, String excelFilePath) throws IOException {
		workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		
		createHeaderRow(sheet);
		
		int rowCount = 0;
		
		br = new BufferedReader(new FileReader(sourceFile));  
		String line = null;  

		CellStyle cellStylePR = sheet.getWorkbook().createCellStyle();
		CellStyle cellStyleVI = sheet.getWorkbook().createCellStyle();
		CellStyle cellStyleEND = sheet.getWorkbook().createCellStyle();
		
		cellStylePR.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		cellStylePR.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyleVI.setFillForegroundColor(IndexedColors.RED.index);
		cellStyleVI.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyleEND.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		cellStyleEND.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		while ((line = br.readLine()) != null)  
		{  
			if(line.contains("STCA")||line.contains("MSAW")||line.contains("APW")) {
				Row row = sheet.createRow(++rowCount);
				if (line.contains("PR")) {
					writeBook(line, row,cellStylePR);
				}
				else if (line.contains("VI")) {
					writeBook(line, row,cellStyleVI);
				}
				else {
					writeBook(line, row,cellStyleEND);
					
				}
			}  
		} 
		
		try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
			workbook.write(outputStream);
		}		
	}

	private void createHeaderRow(Sheet sheet) {
		
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		Font font = sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setColor((short) 5);
		font.setFontHeightInPoints((short) 16);
		cellStyle.setFont(font);
				
		Row row = sheet.createRow(0);
		Cell cellTitle = row.createCell(1);

		cellTitle.setCellStyle(cellStyle);
		cellTitle.setCellValue("Title");
		
		Cell cellAuthor = row.createCell(2);
		cellAuthor.setCellStyle(cellStyle);
		cellAuthor.setCellValue("Author");
		
		Cell cellPrice = row.createCell(3);
		cellPrice.setCellStyle(cellStyle);
		cellPrice.setCellValue("Price");
	}
	
	private void writeBook(String line, Row row, CellStyle cellStyle) {
    		line=line.replaceAll("      ", " ,").trim().replaceAll(" +", " ").replace(' ', ',');
    		
    		//STCA VI, END thieu 2 truong nen chen 2 truong blank vao
    		if((line.contains("STCA")&&line.contains("END"))||line.contains("STCA")&&line.contains("VI")) {
    			int index = 0;
    		    for(int i = 0; i < 9; i++)
    		        index = line.indexOf(",", index+1);
    			line = new StringBuffer(line).insert(index, ",").toString();
    			
    		    for(int i = 0; i < 6; i++)
    		        index = line.indexOf(",", index+1);
    			line = new StringBuffer(line).insert(index, ",").toString();
    		}
    		
    		//dua loai canh bao (VI, PR, END) cua MSAW va APW cung cot voi cac canh bao cua STCA
    		if(line.contains("MSAW")||line.contains("APW"))
    		{
    			int index = 0;
    		    for(int i = 0; i < 7; i++)
    		        index = line.indexOf(",", index+1);
    			line = new StringBuffer(line).insert(index, ",,,,,,,,,").toString();
    		}
    		
//    		write data to workbook
    		Cell cell;
			int i=0;
			for (String infoInLine : line.split(",")) {
				cell = row.createCell(i++);
				cell.setCellValue(infoInLine);
				cell.setCellStyle(cellStyle);
			}

		}
}
