package org.example;


import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


import java.awt.Color;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;

import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ImageIcon;

public class App {
    static XWPFDocument docx;
    static FileOutputStream fos;
    static public JCheckBox alert;

    public static void main(String[] args) throws Exception{

        /* change the jar icon */

        docx = new XWPFDocument();
        String maindir = System.getProperty("user.dir");
        SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd hh mm ss a");
        Calendar now = Calendar.getInstance();
        //File directory = new File(maindir +"\\WordDocument\\");
        File directory = new File(maindir ,"NotesWithScreen");
        if (! directory.exists()){
            directory.mkdir();

        }

        boolean while_loop = true;
        JTextField topicTitle = new JTextField("Name a file");

        JTextArea imagetopic = new JTextArea(4,30);

        JScrollPane scroll = new JScrollPane(imagetopic); //place the JTextArea in a scroll pane

        imagetopic.setLineWrap(true);

        imagetopic.append("This is automated screen snap");
        JCheckBox alert = new JCheckBox("Alert - Note with color");
        JCheckBox snaps = new JCheckBox("Snaps");

        alert.setForeground(Color.RED);


        Object[] message = {topicTitle,scroll,alert,snaps};
        Object[] options = { "Capture Screen", "Exit & Save" };

        ImageIcon icon= new ImageIcon(getImage("images/logo.png"));


        while(while_loop)
        {
            imagetopic.setText(null);
            imagetopic.append("Note: ");
            int n = JOptionPane.showOptionDialog(new JFrame(),
                    message, "Notes with Screen snaps",
                    //JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null,
                    JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, icon,
                    options, options[0]);
            //snaps.requestFocus(); //bring back the window
            //System.out.println("n option-->" + n);
            Thread.sleep(100);
            if ( n==0)
            {
                if (snaps.isSelected())
                {
                    Screenandword("Automated snaps",false);
                    for (int i=0;i<20;i++ )
                    {
                        Thread.sleep(100); //.1 sec
                        Screenandword("Automated snaps# " + Integer.toString(i),false);

                    }

                }

                Screenandword(imagetopic.getText(), alert.isSelected());
            }
            else if ( n==1 || n==-1)
            {
                while_loop=false;
                break;


            }

        }
        String filename=topicTitle.getText();
        if (filename.equalsIgnoreCase("Name a file"))
            filename="Word_Document";


        // os independent file paths
        String joinedPath=new File(maindir, "NotesWithScreen").toString();
        joinedPath = new File(joinedPath, filename + formatter.format(now.getTime())).toString() + "_ScreenNotes.docx";
        fos = new FileOutputStream(joinedPath);
        System.out.println("Filename--" + joinedPath);

        // set the authors

        POIXMLProperties xmlProps = docx.getProperties();
        POIXMLProperties.CoreProperties coreProps = xmlProps.getCoreProperties();
        coreProps.setCreator("ScreenSnap");
        coreProps.setTitle("An Automated Tool");

        docx.write(fos);


        try{
            fos.close();
            System.out.println(" File Saved successfully");
        }
        catch  (Exception e )
        {

            System.out.println(" Sorry, could not save the file." + e.getMessage());

        }
        System.exit(0);
    }

    public static void Screenandword(String textforimage,boolean color_text) throws Exception
    {
        // docx = new XWPFDocument();

        XWPFParagraph par = docx.createParagraph();
        XWPFRun run = par.createRun();

        run.setColor("5c304f");

        if (color_text)
            run.setColor("db172a");

        run.setText(textforimage);
        run.setFontSize(10);
        run.addBreak();
        Robot robot = new Robot();
        BufferedImage screenShot = robot.createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
        // convert buffered image to Input Stream
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(screenShot, "jpeg", baos);
        ByteArrayInputStream bis = new ByteArrayInputStream(baos.toByteArray());
        //XWPFDocument.
        run.addPicture(bis, XWPFDocument.PICTURE_TYPE_PNG, "picture", Units.toEMU(600), Units.toEMU(300));
        run.addBreak();
        baos.close();
        bis.close();

    }

    public static void trail(String[] args) throws Exception{
        XWPFDocument docx = new XWPFDocument();
        XWPFParagraph par = docx.createParagraph();
        XWPFRun run = par.createRun();
        run.setText("Automatic Screen snap for Test only");
        run.setFontSize(10);
        run.addBreak();
        Robot robot = new Robot();
        BufferedImage screenShot = robot.createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
        // convert buffered image to Input Stream
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(screenShot, "jpeg", baos);
        ByteArrayInputStream bis = new ByteArrayInputStream(baos.toByteArray());
        // run.addPicture(bis,Document.PICTURE_TYPE_PNG, "3", 0, 0);
        run.addPicture(bis, XWPFDocument.PICTURE_TYPE_JPEG, "picture", Units.toEMU(600), Units.toEMU(300)); // 200x200 pixels
        baos.close();
        run.addBreak();

        // write word doc to file
        String maindir = System.getProperty("user.dir");
        FileOutputStream fos = new FileOutputStream(maindir + "ScreenNotes.docx");
        docx.write(fos);
        bis.close();
        fos.close();
    }

    public static Image getImage(final String pathAndFileName) {
        final URL url = Thread.currentThread().getContextClassLoader().getResource(pathAndFileName);
        return Toolkit.getDefaultToolkit().getImage(url);
        //Use the following code in case need to use your logo and resize
        //return Toolkit.getDefaultToolkit().getImage(url).getScaledInstance( 70, 90,  java.awt.Image.SCALE_SMOOTH ) ;

    }

}
