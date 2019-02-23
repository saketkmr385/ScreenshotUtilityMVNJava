package ScreenShotUtility;
import java.awt.FlowLayout;
import java.awt.Image;
import java.awt.List;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Scanner;
import java.util.TimeZone;
import java.util.concurrent.TimeUnit;
import javax.imageio.ImageIO;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ScreenShotutility2 {
	
	public static int flag;
	public static XWPFDocument docx;
	public static XWPFRun run;
	public static FileOutputStream out;
	private static JDialog d;
	
	
	public static void main(String[] args)
    {
		 
			 String x;
			 
		 
	      System.out.println("Application started! Enter Keyword S to proceed ");
	      
	    	  try {
	    		  
	    		  String WordDocPath = DocxPath();    
	            	   docx = new XWPFDocument();   	            	   
		               run = docx.createParagraph().createRun();
		               out = new FileOutputStream(WordDocPath);
		               
	               for (int Count = 0; Count<=100; Count++)
	               {
	            	   
		               int captureflag = CaptureCommand();
		               if (captureflag == 0)
		               {
		            	   break;
		               }	               
	               
	               }
	               
	               docx.write(out);
		           out.flush();
		           out.close();
	    		  }
	    	  catch (Exception e) 
	    	  {
	              e.printStackTrace();
	          }
	    	  
	    		  
	      
	      System.exit(0);
			 }
	      


	//Method to capture and delete screenshots
	public static void captureScreenShot(XWPFDocument docx, XWPFRun run, FileOutputStream out) throws Exception 
	{
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(new Date());
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        sdf.setTimeZone(TimeZone.getTimeZone("EST"));
        String Timestamp = sdf.format(calendar.getTime());

		//Creates a temp folder at User home location
		File dir = new File(System.getProperty("user.home") +"//ScreenshotTemp//");
		{
	    if (!dir.exists()) dir.mkdirs();
		}
		
		
        String screenshot_name = System.currentTimeMillis() + ".png";
        BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));  
        File file = new File(System.getProperty("user.home") +"//ScreenshotTemp//" + screenshot_name);
        ImageIO.write(image, "png", file);
        InputStream pic = new FileInputStream(System.getProperty("user.home")+"//ScreenshotTemp//" + screenshot_name);
        run.addBreak();
        run.setText(Timestamp);
        run.addPicture(pic, XWPFDocument.PICTURE_TYPE_PNG, screenshot_name, Units.toEMU(500), Units.toEMU(350));
        pic.close();
        file.delete();
        
    }
	
	//Method to choose the doc file to paste screenshots	
	public static String DocxPath () throws IOException
	{
		
		String DocxPath = "";
		String DocxJavaPath = "";
		
		String LastUserSelectionpath = getLastSelectionPath();
		
		
	JFileChooser fileChooser = new JFileChooser();
	 fileChooser.setDialogTitle("Word Doc Path");
	 
	 if(LastUserSelectionpath != null && !LastUserSelectionpath.isEmpty())
		 fileChooser.setCurrentDirectory(new File(LastUserSelectionpath)); 
	 
	 
	 else
		 fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
		 
	 fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
	 
	 fileChooser.setFileFilter(new FileNameExtensionFilter("MS Office Documents", "docx"));
	 
    int result = fileChooser.showOpenDialog(null);
    
        if (result == JFileChooser.APPROVE_OPTION) {
        File selectedFile = fileChooser.getSelectedFile();
        System.out.println("Selected file: " + selectedFile.getAbsolutePath());
        
        DocxPath = selectedFile.getAbsolutePath();
        String EscSeq = "\"";
        String DoubleEscSeq = "//";        
        DocxJavaPath = DocxPath.replaceAll(EscSeq,DoubleEscSeq);
        System.out.println(DocxJavaPath);
        writePathToNotepad(DocxJavaPath);     
        
        
    } 
    
    return DocxJavaPath;
	}
	
	
	
	//method to generate a JFRame to capture screenshot
	public  static int  CaptureCommand(){  		
		
		flag = 1;
		JFrame Frame = new JFrame();
		
		ImageIcon logoicon = new ImageIcon("download.jfif");
		Image logo = logoicon.getImage();
		Frame.setIconImage(logo);
		
		d = new JDialog(Frame , "Screen Capture Utility", true);
		d.setLayout( new FlowLayout());
		
		JButton button = new JButton("Capture");
		JButton Close = new JButton ("Close");
				
		button.addActionListener ( new ActionListener()  
        {  
            public void actionPerformed( ActionEvent e )  
            {  
            	try {
					captureScreenShot(docx, run, out);
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
                ScreenShotutility2.d.setVisible(false);  
            }  
        });
		
		Close.addActionListener ( new ActionListener()  
        {  
            public void actionPerformed( ActionEvent e2 )  
            {  
            	try {
            		flag = 0;
            		d.dispose();
            		
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
                ScreenShotutility2.d.setVisible(false);  
            }  
        });
		
		
		
		d.add(button);
		d.add(Close);
		d.setSize(300,100);    
        d.setVisible(true);	
		
		
		return flag;
				
		
	    } 
	
	public static String getLastSelectionPath() throws IOException
	{
		File dir = new File(System.getProperty("user.home") +"//ScreenshotTemp//");
		{
	    if (!dir.exists()) dir.mkdirs();
		}
		
		File f = new File(System.getProperty("user.home") +"//ScreenshotTemp//temp.txt");
		if(!f.exists()){
		  f.createNewFile();
		}
		
		BufferedReader reader = new BufferedReader(new FileReader(System.getProperty("user.home") +"//ScreenshotTemp//temp.txt"));
		String Path;
		String LastUserPath="";
		  //as long as there are lines in the file, print them
		  while((Path = reader.readLine()) != null)
		  { 
		    System.out.println(Path);
		    LastUserPath = Path;
		    
		  }
		  
		  return LastUserPath;
	}
	
	public static void writePathToNotepad(String path) throws IOException
	{
		File file = new File(System.getProperty("user.home") +"//ScreenshotTemp//temp.txt");
		BufferedWriter wr = new BufferedWriter(new FileWriter(file));
		wr.write(path);
		wr.close();
	}
	
	
}
