package practice_poiPDF;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import javax.swing.plaf.synth.SynthSpinnerUI;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

public class CreatePdf {
	
	public static final String RESULT = "test.pdf";
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		XWPFDocument document = new  XWPFDocument();
		
		try{
			new CreatePdf().createPdf(RESULT);
			System.out.println("File Created");
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
	}
	
	public void createPdf(String filename)
			throws DocumentException, IOException {
		        Document document = new Document();

		        PdfWriter.getInstance(document, new FileOutputStream(filename));
		        document.open();

		        String para = "Hello! this is the sample text file used to test the pdf generation";
		        document.add(new Paragraph(para));

		        document.close();
		    }
	
}
