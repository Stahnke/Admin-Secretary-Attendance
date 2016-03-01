import java.awt.FlowLayout;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.util.Collections;
import java.util.List;

import javax.swing.BoxLayout;
import javax.swing.JFrame;
import javax.swing.JTextArea;
 

public class DragAndDrop extends JFrame{

	private ExcelReader excelReader = new ExcelReader();
	
	public static void main(String args[]){
		
		new DragAndDrop().setVisible(true);
		
	}
	
	private DragAndDrop(){
		
		super("Drag And Drop");
		setSize(400,400);
		setResizable(false);
		
		setLayout(new BoxLayout(this.getContentPane(), BoxLayout.Y_AXIS));
		JTextArea headline = new JTextArea();
		add(headline);
		headline.setText("Please drag and drop files here:");
		headline.setEditable(false);
		JTextArea myPanel = new JTextArea();
		add(myPanel);
		myPanel.setDropTarget(new DropTarget() {
		    public synchronized void drop(DropTargetDropEvent evt) {
		        try {
		            evt.acceptDrop(DnDConstants.ACTION_COPY);		        
		            List<File> droppedFiles = (List<File>)
		                evt.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
		            	Collections.sort(droppedFiles);
		            	excelReader.getExcelSheet(droppedFiles);
		            for (File file : droppedFiles) {
		                System.out.println(file.getName());
		                myPanel.append(file.getName());
		                myPanel.append("\n");	                
		            }
		        } catch (Exception ex) {
		            ex.printStackTrace();
		        }
		    }
		});
		
		
	}
}
