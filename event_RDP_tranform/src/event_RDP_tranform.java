import java.awt.EventQueue;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;
import java.util.Scanner;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.swing.JMenuBar;
import javax.swing.JMenu;
import javax.swing.JMenuItem;
import javax.swing.JTable;
import javax.swing.RowFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableRowSorter;
import javax.swing.JScrollPane;
import javax.swing.JCheckBox;
import javax.swing.JComponent;
import javax.swing.AbstractButton;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;

import java.awt.GridLayout;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Component;

import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import java.awt.BorderLayout;
import javax.swing.JPanel;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import javax.swing.JTextPane;
import javax.swing.JTextArea;
import javax.swing.JToggleButton;
import javax.swing.event.ChangeListener;
import javax.swing.event.ChangeEvent;
import java.awt.event.ItemListener;
import java.awt.event.ItemEvent;

public class event_RDP_tranform {

	private JFrame frmRdpEventTxtexcel;
	private JTable table;
	protected DefaultTableModel dm;
	String[] alertTypes = {"MSAW", "STCA", "APW","PR", "VI", "END"};
	private JTextArea textArea;
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
		frmRdpEventTxtexcel.setBounds(100, 100, 501, 384);
		frmRdpEventTxtexcel.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		//create folder to save data export if not exist
		try {
			Files.createDirectories(Paths.get(System.getProperty("user.dir")+"/data export"));
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		
		JScrollPane scrollPane = new JScrollPane();
		
		
		
		JMenuBar menuBar = new JMenuBar();
		frmRdpEventTxtexcel.setJMenuBar(menuBar);
		frmRdpEventTxtexcel.getContentPane().setLayout(new BorderLayout(0, 0));
		
		JPanel mainPanel = new JPanel();
		frmRdpEventTxtexcel.getContentPane().add(mainPanel, BorderLayout.NORTH);
		mainPanel.setLayout(new BorderLayout(0, 0));
		
		JPanel topPanel = new JPanel();
		mainPanel.add(topPanel, BorderLayout.NORTH);
		topPanel.setLayout(new BoxLayout(topPanel, BoxLayout.X_AXIS));
		
		String[] food = {"MSAW", "STCA", "APW"};

		JCheckBox[] boxes = new JCheckBox[food.length]; //  Each checkbox will get a name of food from food array.  
		
		for(int i = 0; i < boxes.length; i++)
		{
			boxes[i] = new JCheckBox(food[i]);
			boxes[i].setSelected(true);
		
			topPanel.add(boxes[i]);	
		}
		
		JButton btnNewButton = new JButton("Filter");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				TableRowSorter<DefaultTableModel> tr=new TableRowSorter<DefaultTableModel>(dm);
				table.setRowSorter(tr);
				
				List<RowFilter<Object,Object>> filters = new ArrayList<RowFilter<Object,Object>>(2);
				
				for(JCheckBox box:boxes)
				if(box.isSelected())
				   filters.add(RowFilter.regexFilter(box.getText()));
				   RowFilter<Object,Object> fooBarFilter = RowFilter.orFilter(filters);
				   tr.setRowFilter(fooBarFilter);
			}
		});
		topPanel.add(btnNewButton);
		
		JToggleButton tglbtnColorFill = new JToggleButton("Fill Color");
		tglbtnColorFill.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
table.repaint();
			}
		});
		
		
		topPanel.add(tglbtnColorFill);
		frmRdpEventTxtexcel.getContentPane().add(scrollPane);
		
		JPanel bottomPanel = new JPanel();
		frmRdpEventTxtexcel.getContentPane().add(bottomPanel, BorderLayout.SOUTH);
		bottomPanel.setLayout(new BorderLayout(0, 0));
		
		JPanel panel = new JPanel();
		bottomPanel.add(panel, BorderLayout.NORTH);
		panel.setLayout(new BorderLayout(0, 0));
		
		textArea = new JTextArea();
		textArea.setToolTipText("Error display, if there are any text in this display, please contact author because there are some exported data error!");
		textArea.setLineWrap(true);
		textArea.setWrapStyleWord(true);
		textArea.setRows(3);
		textArea.setEditable(false);
		JScrollPane scroll = new JScrollPane(textArea);
		panel.add(scroll, BorderLayout.NORTH);
		JMenu mnFile = new JMenu("File");
		menuBar.add(mnFile);
		
		JMenuItem mntmOpen = new JMenuItem("Open");
		mntmOpen.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setMultiSelectionEnabled(true);
				int returnValue = fileChooser.showOpenDialog(null);
				if (returnValue == JFileChooser.APPROVE_OPTION) {
					File[] selectedFile = fileChooser.getSelectedFiles();
					System.out.println(selectedFile);

					// copy content nhieu file event txt lai thanh 1 file txt roi moi chuyen file
					// txt do qua excel
					new File("NewFile.txt").delete();
					File allContentFile = new File("NewFile.txt");
					File[] fileLocations = new File[selectedFile.length];

					for (int i = 0; i < selectedFile.length; i++) {
						fileLocations[i] = new File(selectedFile[i].getAbsolutePath());
					}
					mergeFiles(fileLocations, allContentFile);
					
					for (int i = 0; i < selectedFile.length; i++) {
						try {
							populate(allContentFile.getAbsolutePath());
						} catch (IOException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}

					}
					JOptionPane.showMessageDialog(new JFrame(), "Transformation completed\nData loaded on table", "Dialog",
					        JOptionPane.INFORMATION_MESSAGE);
				}
			}

			private void populate(String absolutePath) throws IOException {
				// TODO Auto-generated method stub
				BufferedReader br = new BufferedReader(new FileReader(absolutePath));
				String line = null;

				List<String> headerRow = new ArrayList<String>();	
				File fileDir = new File("header.txt");
				
				BufferedReader in = new BufferedReader(
				   new InputStreamReader(
		                      new FileInputStream(fileDir), "UTF8"));
				        
				String str;
				      
				while ((str = in.readLine()) != null) {
					headerRow.add(str);
				}
				dm=new DefaultTableModel(
						null,
						headerRow.toArray()
					);
				int rowCount=0;
				while ((line = br.readLine()) != null) {
					
					if (line.contains("STCA") || line.contains("MSAW") || line.contains("APW")) {
						line = lineTablingCorrection(line);
						
						line=++rowCount+","+line;
						String[] rowdata= line.split(",");
						dm.addRow(rowdata);
					}
				}

				table =(JTable) createData(dm);
				scrollPane.setViewportView(table);
				table.getParent().addComponentListener(new ComponentAdapter() {
				    @Override
				    public void componentResized(final ComponentEvent e) {
				        if (table.getPreferredSize().width < table.getParent().getWidth()) {
				        	table.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);
				        } else {
				        	table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
				        }
				    }
				});
				
				try {
					br.close();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}

			private JComponent createData(DefaultTableModel model)
			{
				
			
			
				
				
				
				JTable table = new JTable( model )
				{
					
					public Component prepareRenderer(TableCellRenderer renderer, int row, int column)
					{
						Component c = super.prepareRenderer(renderer, row, column);

						//  Color row based on a cell value

						if (!isRowSelected(row))
						{
							c.setBackground(getBackground());
							int modelRow = convertRowIndexToModel(row);
							if (tglbtnColorFill.isSelected()) {
								String type = (String) getModel().getValueAt(modelRow, 17);
								if ("VI".equals(type))
									c.setBackground(Color.RED);
								if ("PR".equals(type))
									c.setBackground(Color.YELLOW);
							}
						}

						return c;
					}
				};
				
				
				table.setPreferredScrollableViewportSize(table.getPreferredSize());
				table.changeSelection(0, 0, false, false);
		        table.setAutoCreateRowSorter(true);
				return table;
			}
		});
		mnFile.add(mntmOpen);
		
		JMenu mnExport = new JMenu("Export");
		menuBar.add(mnExport);
		
		JMenuItem mntmNewMenuItem = new JMenuItem("select event file (many txt -> 1 xls)");
		mntmNewMenuItem.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setMultiSelectionEnabled(true);
				int returnValue = fileChooser.showOpenDialog(null);
				if (returnValue == JFileChooser.APPROVE_OPTION) {
					File[] selectedFile = fileChooser.getSelectedFiles();
					System.out.println(selectedFile);

					// copy content nhieu file event txt lai thanh 1 file txt roi moi chuyen file
					// txt do qua excel
					new File("NewFile.txt").delete();
					File allContentFile = new File("NewFile.txt");
					File[] fileLocations = new File[selectedFile.length];

					for (int i = 0; i < selectedFile.length; i++) {
						fileLocations[i] = new File(selectedFile[i].getAbsolutePath());
					}
					mergeFiles(fileLocations, allContentFile);
					
					//form for user choose location where save the xls file
					JFileChooser fileChooser1 = new JFileChooser(System.getProperty("user.dir") + "\\data export\\");
					fileChooser.setDialogTitle("Specify a file to save");   	
					int userSelection = fileChooser1.showSaveDialog(frmRdpEventTxtexcel);
					File fileToSave=null;
					if (userSelection == JFileChooser.APPROVE_OPTION) {
					    fileToSave = fileChooser1.getSelectedFile();
					    System.out.println("Save as file: " + fileToSave.getAbsolutePath());
					}
					else {
						return;
					}
			        
					String excelFilePath = fileToSave.getAbsolutePath()+".xls";
					try {
						writeExcelColorFill(allContentFile.getAbsolutePath(),
								excelFilePath);
						JOptionPane.showMessageDialog(new JFrame(), "Transformation completed", "Dialog",
						        JOptionPane.INFORMATION_MESSAGE);
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			}
		
			
		});
		
		mnExport.add(mntmNewMenuItem);
		
		JMenuItem mntmSelectEventFile = new JMenuItem("select event file (1 txt -> 1 xls)");
		mntmSelectEventFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setMultiSelectionEnabled(true);
				int returnValue = fileChooser.showOpenDialog(null);
				if (returnValue == JFileChooser.APPROVE_OPTION) {
					File[] selectedFile = fileChooser.getSelectedFiles();
					System.out.println(selectedFile);

					try {
						for (int i = 0; i < selectedFile.length; i++) {
							// dataTransform(selectedFile[i].getParent(), selectedFile[i].getName());
							String excelFilePath = System.getProperty("user.dir") + "\\data export\\" + selectedFile[i].getName() + ".xls";
							writeExcelColorFill(selectedFile[i].getAbsolutePath(),
									excelFilePath);
						}
						JOptionPane.showMessageDialog(new JFrame(), "Transformation completed\nFile exported in data export folder", "Dialog",
						        JOptionPane.INFORMATION_MESSAGE);
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			}
		
			
		});
		mnExport.add(mntmSelectEventFile);
		
		JMenuItem mntmPirintableExport = new JMenuItem("Printable export");
		mntmPirintableExport.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 boolean[] alertTypeChoosen= chooseAlertTypeToExport_Window();

				 //neu user chon 1 checkbox trong 3 options MSAW STCA APW va 1 checkbox trong 3 options PR VI END o chooseAlertTypeToExport_Window thi thuc hien xuat file theo option duoc chon 
					if (minimumConditionChoosen(alertTypeChoosen)) {
						JFileChooser fileChooser = new JFileChooser();
						fileChooser.setMultiSelectionEnabled(true);
						int returnValue = fileChooser.showOpenDialog(null);
						if (returnValue == JFileChooser.APPROVE_OPTION) {
							File[] selectedFile = fileChooser.getSelectedFiles();
							System.out.println(selectedFile);

							// copy content nhieu file event txt lai thanh 1 file txt roi moi chuyen file
							// txt do qua excel
							new File("NewFile.txt").delete();
							File allContentFile = new File("NewFile.txt");
							File[] fileLocations = new File[selectedFile.length];

							for (int i = 0; i < selectedFile.length; i++) {
								fileLocations[i] = new File(selectedFile[i].getAbsolutePath());
							}
							mergeFiles(fileLocations, allContentFile);

							//form for user choose location where save the xls file
							JFileChooser fileChooser1 = new JFileChooser(
									System.getProperty("user.dir") + "\\data export\\");
							fileChooser.setDialogTitle("Specify a file to save");
							int userSelection = fileChooser1.showSaveDialog(frmRdpEventTxtexcel);
							File fileToSave = null;
							if (userSelection == JFileChooser.APPROVE_OPTION) {
								fileToSave = fileChooser1.getSelectedFile();
								System.out.println("Save as file: " + fileToSave.getAbsolutePath());
							} else {
								return;
							}

							String excelFilePath = fileToSave.getAbsolutePath() + ".xls";
							try {
								writeExcelPrintable(allContentFile.getAbsolutePath(), excelFilePath,alertTypeChoosen,selectedFile);
								JOptionPane.showMessageDialog(new JFrame(), "Transformation completed", "Dialog",
										JOptionPane.INFORMATION_MESSAGE);
							} catch (IOException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
						} 
					}
				

			}

			private boolean minimumConditionChoosen(boolean[] alertTypeChoosen) {
				return containsTrue(Arrays.copyOfRange(alertTypeChoosen, 0, alertTypeChoosen.length/2))&&containsTrue(Arrays.copyOfRange(alertTypeChoosen, alertTypeChoosen.length/2, alertTypeChoosen.length));
				
			}

			private boolean containsTrue(boolean[] array){

			    for(boolean val : array){
			        if(val)
			            return true;
			    }

			    return false;
			}

			private void writeExcelPrintable(String sourceFile, String excelFilePath, boolean[] alertTypeChoosen, File[] selectedFile) throws IOException {
				Workbook workbook = new HSSFWorkbook();
				
				
				//tinh so sheet tao ra: loai canh bao * loai PR VI END * so file event moi ngay
				List<String> alertTypeMsawStcaApwChoosen=new ArrayList<String>();
				for (int i = 0; i < alertTypeChoosen.length/2; i++) {
					if (alertTypeChoosen[i]) {
						alertTypeMsawStcaApwChoosen.add(alertTypes[i]);
					}
				}
				List<String> alertTypePrViEndChoosen=new ArrayList<String>();
				for (int i = alertTypeChoosen.length/2; i < alertTypeChoosen.length; i++) {
					if (alertTypeChoosen[i]) {
						alertTypePrViEndChoosen.add(alertTypes[i]);
					}
				}
				
				
					for (int i = 0; i < selectedFile.length; i++) {
						for (int j = 0; j < 3; j++) {
							for (int k = 0; k < 3; k++) {
							if (alertTypeChoosen[j]&&alertTypeChoosen[k+3]) {
								
								workbook.createSheet(selectedFile[i].getName() +" "+ alertTypes[j] +" "+ alertTypes[k + 3]);
							}
							}
						}
					}
					
					
				ArrayList<Sheet> sheets = new ArrayList<Sheet>();
				for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
					sheets.add(workbook.getSheetAt(i));  
				      //.... code to print the sheet's values here
				}
				for (Sheet sheet:sheets) {
					//table header
					createHeaderRow(sheet);
					//sheet header
					Header header = sheet.getHeader();
					header.setCenter("Ngay "+sheet.getSheetName().split(" ")[0]+", "+sheet.getSheetName().split(" ")[1]+" "+sheet.getSheetName().split(" ")[2]);
					sheet.setRepeatingRows(CellRangeAddress.valueOf("1:1"));
//					header.setLeft("Left Header");
//					header.setRight(HSSFHeader.font("Stencil-Normal", "Italic")
//					+ HSSFHeader.fontSize((short) 10) + "Right Header");
					
					sheet.getPrintSetup().setScale((short)55);
					sheet.getPrintSetup().setLandscape(true);
					
				}
				
				int[] rowCount = new int[workbook.getNumberOfSheets()];
				Arrays.fill(rowCount, 0);
				
				BufferedReader br = new BufferedReader(new FileReader(sourceFile));
				String line = null;

				CellStyle cellStyleEND = workbook.createCellStyle();
				cellStyleEND.setFillForegroundColor(IndexedColors.WHITE.getIndex());
				cellStyleEND.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				cellStyleEND.setBorderLeft(BorderStyle.THIN);
				cellStyleEND.setBorderBottom(BorderStyle.THIN);
				cellStyleEND.setBorderRight(BorderStyle.THIN);
				cellStyleEND.setBorderTop(BorderStyle.THIN);
				

				while ((line = br.readLine()) != null) {
					if (stringContainAnElementFromListOfString(line,alertTypeMsawStcaApwChoosen)&&stringContainAnElementFromListOfString(line, alertTypePrViEndChoosen)) {
						
						if (line.contains("PR")) {
							
							int sheetIndex= workbook.getSheetIndex(line.split(" ")[1].replace('/', '_')+"_h "+line.split(" ")[7]+" PR");
							Row row = sheets.get(sheetIndex).createRow(++rowCount[sheetIndex]);
							writeBook(line, row, cellStyleEND);
							
						} else if (line.contains("VI")) {
							int sheetIndex= workbook.getSheetIndex(line.split(" ")[1].replace('/', '_')+"_h "+line.split(" ")[7]+" VI");
							Row row = sheets.get(sheetIndex).createRow(++rowCount[sheetIndex]);
							writeBook(line, row, cellStyleEND);
						} else {
							int sheetIndex= workbook.getSheetIndex(line.split(" ")[1].replace('/', '_')+"_h "+line.split(" ")[7]+" END");
							Row row = sheets.get(sheetIndex).createRow(++rowCount[sheetIndex]);
							writeBook(line, row, cellStyleEND);

						}
					}
				}

				try {
					br.close();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
					workbook.write(outputStream);
					
				}
			 catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				JOptionPane.showMessageDialog(new JFrame(), "Excel file is opening, please close it!", "Dialog",
				        JOptionPane.ERROR_MESSAGE);
				System.exit(0);
			}
				
			}



			/**
			 * @return 
			 * 
			 */
			private boolean[] chooseAlertTypeToExport_Window() {
				
				    JCheckBox[] check = new JCheckBox[alertTypes.length];

				    for(int i = 0; i < alertTypes.length; i++)
				        check[i] = new JCheckBox(alertTypes[i]);    

				    //set default checkbox is checked
				    check[0].setSelected(true);
				    check[1].setSelected(true);
				    check[4].setSelected(true);
				    
				    
				    boolean[] ret = new boolean[alertTypes.length];     

				    int answer = JOptionPane.showConfirmDialog(null, new Object[]{"Choose alerts type for exporting:", check}, "Choose alerts type" , JOptionPane.OK_CANCEL_OPTION);

				    if(answer == JOptionPane.OK_OPTION)
				    {
				        for(int i = 0;i < alertTypes.length ; i++)
				            ret[i] = check[i].isSelected();

				    }else if(answer == JOptionPane.CANCEL_OPTION || answer == JOptionPane.ERROR_MESSAGE)
				    {
				        for(int i = 0; i < alertTypes.length; i++)
				            ret[i] = false;
				    }

				    return ret;
			}
		});
		mnExport.add(mntmPirintableExport);
		
		JMenu mnHelp = new JMenu("Help");
		menuBar.add(mnHelp);
		
		JMenuItem mntmAbout = new JMenuItem("About");
		mntmAbout.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JOptionPane.showMessageDialog(new JFrame(), "Phan mem dinh dang cac canh bao cua he thong RDP Aircon 2100 giup viec doc, trich xuat cac canh bao theo yeu cau de dang hon\nVersion 1.0", "About",
				        JOptionPane.INFORMATION_MESSAGE);
			
			}
		});
		mnHelp.add(mntmAbout);
	}

	protected void writeExcelColorFill(String sourceFile, String excelFilePath) throws IOException {
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();

		createHeaderRow(sheet);

		int rowCount = 0;

		BufferedReader br = new BufferedReader(new FileReader(sourceFile));
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
		cellStylePR.setBorderLeft(BorderStyle.THIN);
		cellStylePR.setBorderBottom(BorderStyle.THIN);
		cellStylePR.setBorderRight(BorderStyle.THIN);
		cellStylePR.setBorderTop(BorderStyle.THIN);
		cellStyleVI.setBorderLeft(BorderStyle.THIN);
		cellStyleVI.setBorderBottom(BorderStyle.THIN);
		cellStyleVI.setBorderRight(BorderStyle.THIN);
		cellStyleVI.setBorderTop(BorderStyle.THIN);
		cellStyleEND.setBorderLeft(BorderStyle.THIN);
		cellStyleEND.setBorderBottom(BorderStyle.THIN);
		cellStyleEND.setBorderRight(BorderStyle.THIN);
		cellStyleEND.setBorderTop(BorderStyle.THIN);

		while ((line = br.readLine()) != null) {
			if (line.contains("STCA") || line.contains("MSAW") || line.contains("APW")) {
				Row row = sheet.createRow(++rowCount);
				if (line.contains("PR")) {
					writeBook(line, row, cellStylePR);
				} else if (line.contains("VI")) {
					writeBook(line, row, cellStyleVI);
				} else {
					writeBook(line, row, cellStyleEND);

				}
			}
		}

		try {
			br.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
			workbook.write(outputStream);
			
		}
	 catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
		JOptionPane.showMessageDialog(new JFrame(), "Excel file is opening, please close it!", "Dialog",
		        JOptionPane.INFORMATION_MESSAGE);
		System.exit(0);
	}
	}

	private void createHeaderRow(Sheet sheet) {

		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		Font font = sheet.getWorkbook().createFont();
		font.setBold(true);
//		font.setColor((short) 5);
//		font.setFontHeightInPoints((short) 16);
		cellStyle.setFont(font);
		
		
		
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		
		List<String> lines = new ArrayList<String>();	
		try {
			File fileDir = new File("header.txt");
				
			BufferedReader in = new BufferedReader(
			   new InputStreamReader(
	                      new FileInputStream(fileDir), "UTF8"));
			        
			String str;
			      
			while ((str = in.readLine()) != null) {
				lines.add(str);
			}
			        
	                in.close();
		    } 
		    catch (UnsupportedEncodingException e) 
		    {
				System.out.println(e.getMessage());
		    } 
		    catch (IOException e) 
		    {
				System.out.println(e.getMessage());
		    }
		    catch (Exception e)
		    {
				System.out.println(e.getMessage());
		    }
		
		

		String[] toppings = lines.toArray(new String[0]);



		Row row = sheet.createRow(0);
		row.setHeight((short) 1800);
		Cell cellTitle;
		cellStyle.setWrapText(true);
		for (int i = 0; i < toppings.length; i++) {
			cellTitle = row.createCell(i);
			cellTitle.setCellStyle(cellStyle);
			cellTitle.setCellValue(toppings[i]);
			sheet.setColumnWidth(i, 3000);
			
		}

//		Cell cellAuthor = row.createCell(2);
//		cellAuthor.setCellStyle(cellStyle);
//		cellAuthor.setCellValue("Author");
//
//		Cell cellPrice = row.createCell(3);
//		cellPrice.setCellStyle(cellStyle);
//		cellPrice.setCellValue("Price");
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

	private void writeBook(String line, Row row, CellStyle cellStyle) {
		line = lineTablingCorrection(line);

		// write data to workbook
		Cell cell;
		//STT
		cell = row.createCell(0);
		cell.setCellValue(row.getRowNum());
		cell.setCellStyle(cellStyle);
		//data
		int i = 1;
		for (String infoInLine : line.split(",")) {
			cell = row.createCell(i++);
			cell.setCellValue(infoInLine);
			cell.setCellStyle(cellStyle);
		}

	}

	private String lineTablingCorrection(String line) {
		line = line.replaceAll("      ", " ,").replaceAll("( +)"," ").trim().replace(", " , ",").replace(' ', ',');
		
//		try {
//		    Files.write(Paths.get("myfile.txt"), (line+"\n").getBytes(), StandardOpenOption.APPEND);
//		}catch (IOException e) {
//		    //exception handling left as an exercise for the reader
//		}
		
		// STCA VI, END thieu 2 truong nen chen 2 truong blank vao
		if ((line.contains("STCA") && line.contains("END")) || line.contains("STCA") && line.contains("VI")) {
			int index = 0;
			for (int i = 0; i < 9; i++)
				index = line.indexOf(",", index + 1);
			line = new StringBuffer(line).insert(index, ",").toString();

			for (int i = 0; i < 6; i++)
				index = line.indexOf(",", index + 1);
			line = new StringBuffer(line).insert(index, ",").toString();
		}

		// dua loai canh bao (VI, PR, END) cua MSAW va APW cung cot voi cac canh bao cua
		// STCA
		if (line.contains("MSAW") || line.contains("APW")) {
			int index = 0;
			for (int i = 0; i < 7; i++)
				index = line.indexOf(",", index + 1);
			line = new StringBuffer(line).insert(index, ",,,,,,,,,").toString();
		}
		
		//neu PR VI END ko nam dung cot chua, cot 17, xuat report error ra textArea o duoi cua so
		if (!stringContainAnElementFromListOfString(line.split(",")[16], Arrays.asList("PR", "VI", "END"))) {
			textArea.append(new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime())+"\t"+ line+"\n");
		}
		
		return line;
	}

	private boolean stringContainAnElementFromListOfString(String line,
						List<String> list) {
	
					   String string = line;
	
					    boolean match = false;
					    for (String s : list) {
					       if(string.contains(s)){
					           match = true;
					           break;
					       }
					    }
	//				   System.out.println(match);
					return match;
				}
}
