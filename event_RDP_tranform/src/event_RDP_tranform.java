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
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

public class event_RDP_tranform {

	private JFrame frmRdpEventTxtexcel;
	private JTable table;
	protected DefaultTableModel dm;

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
		frmRdpEventTxtexcel.getContentPane().add(scrollPane);
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
					JOptionPane.showMessageDialog(new JFrame(), "Transformation completed\nFile exported in data export folder", "Dialog",
					        JOptionPane.INFORMATION_MESSAGE);
				}
			}

			private void populate(String absolutePath) throws IOException {
				// TODO Auto-generated method stub
				BufferedReader br = new BufferedReader(new FileReader(absolutePath));
				String line = null;

				dm=new DefaultTableModel(
						null,
						new String[] {
							"New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column", "New column"
						}
					);
				while ((line = br.readLine()) != null) {
					
					if (line.contains("STCA") || line.contains("MSAW") || line.contains("APW")) {
						line = lineTablingCorrection(line);
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
							String type = (String)getModel().getValueAt(modelRow, 16);
							if ("VI".equals(type)) c.setBackground(Color.RED);
							if ("PR".equals(type)) c.setBackground(Color.YELLOW);
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
		
		JMenuItem menuItem = new JMenuItem("New menu item");
		mnFile.add(menuItem);
		
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
						writeExcel(allContentFile.getAbsolutePath(),
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
							writeExcel(selectedFile[i].getAbsolutePath(),
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
				 boolean[] alertType= chooseAlertTypeToExport_Window();

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
							writeExcel(allContentFile.getAbsolutePath(),
									excelFilePath);
							JOptionPane.showMessageDialog(new JFrame(), "Transformation completed", "Dialog",
							        JOptionPane.INFORMATION_MESSAGE);
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
					}
				

			}

			/**
			 * @return 
			 * 
			 */
			private boolean[] chooseAlertTypeToExport_Window() {
				String[] genres = {"MSAW", "STCA", "APW","PR", "VI", "END"};
				    JCheckBox[] check = new JCheckBox[genres.length];

				    for(int i = 0; i < genres.length; i++)
				        check[i] = new JCheckBox(genres[i]);    

				    boolean[] ret = new boolean[genres.length];     

				    int answer = JOptionPane.showConfirmDialog(null, new Object[]{"Choose gernes:", check}, "Genres" , JOptionPane.OK_CANCEL_OPTION);

				    if(answer == JOptionPane.OK_OPTION)
				    {
				        for(int i = 0;i < genres.length ; i++)
				            ret[i] = check[i].isSelected();

				    }else if(answer == JOptionPane.CANCEL_OPTION || answer == JOptionPane.ERROR_MESSAGE)
				    {
				        for(int i = 0; i < genres.length; i++)
				            ret[i] = false;
				    }

				    return ret;
			}
		});
		mnExport.add(mntmPirintableExport);
		
		JMenu mnHelp = new JMenu("Help");
		menuBar.add(mnHelp);
		
		JMenuItem mntmAbout = new JMenuItem("About");
		mnHelp.add(mntmAbout);
	}

	protected void writeExcel(String sourceFile, String excelFilePath) throws IOException {
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
		int i = 0;
		for (String infoInLine : line.split(",")) {
			cell = row.createCell(i++);
			cell.setCellValue(infoInLine);
			cell.setCellStyle(cellStyle);
		}

	}

	private String lineTablingCorrection(String line) {
		line = line.replaceAll("      ", " ,").trim().replaceAll(" +", " ").replace(' ', ',');

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
		return line;
	}
}
