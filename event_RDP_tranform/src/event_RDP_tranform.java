import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JButton;
import java.awt.BorderLayout;
import java.awt.FlowLayout;
import javax.swing.JFileChooser;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.Scanner;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.awt.event.ActionEvent;

public class event_RDP_tranform {

	private JFrame frmRdpEventTxtexcel;

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
		
		JButton btnOpen = new JButton("select event file");
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
							new File(selectedFile[i].getAbsolutePath()+".csv").delete();
							dataTransform(selectedFile[i].getParent(), selectedFile[i].getName());
						}
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				    }
			}

			private void dataTransform(String parentFolder, String fileName) throws IOException {
				// TODO Auto-generated method stub
				final Path path = Paths.get(parentFolder);
			    final Path txt = path.resolve(fileName);
			    final Path csv = path.resolve(fileName+".csv");
			    final Charset utf8 = Charset.forName("UTF-8");
			    try (
			            final Scanner scanner = new Scanner(Files.newBufferedReader(txt, utf8));
			            final PrintWriter pw = new PrintWriter(Files.newBufferedWriter(csv, utf8, StandardOpenOption.CREATE_NEW))) {
			        while (scanner.hasNextLine()) {
			        	String line=scanner.nextLine();
			        	if(line.contains("STCA")||line.contains("MSAW")||line.contains("APW")) {
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
			        		
			        		//sua loi callsign 0647 bi excel doc thanh 647
			        		int index = 0;
		        		    for(int i = 0; i < 4; i++)
		        		        index = line.indexOf(",", index+1);
		        			line = new StringBuffer(line).insert(index+1, "\"=\"\"").toString();
		        			index = line.indexOf(",", index+1);
		        			line = new StringBuffer(line).insert(index, "\"\"\"").toString();
		        			
		        			for(int i = 0; i < 6; i++)
		        		        index = line.indexOf(",", index+1);
		        			line = new StringBuffer(line).insert(index+1, "\"=\"\"").toString();
		        			index = line.indexOf(",", index+1);
		        			line = new StringBuffer(line).insert(index, "\"\"\"").toString();
			        		pw.println(line);
			        	}
			        }
			    }
			}
		});
		frmRdpEventTxtexcel.getContentPane().add(btnOpen);
	}

}
