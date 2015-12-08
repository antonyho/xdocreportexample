package net.antonyho.xdocreportexample;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;

import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.images.FileImageProvider;
import fr.opensagres.xdocreport.document.images.IImageProvider;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;
import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class AppTest 
    extends TestCase
{
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( AppTest.class );
    }

    /**
     * Rigourous Test :-)
     * @throws IOException 
     * @throws XDocReportException 
     */
    public void testApp() throws IOException, XDocReportException
    {
    	final URL docResource = this.getClass().getResource("/template1.docx");
    	InputStream is = docResource.openStream();
    	// we use Freemarker here
    	IXDocReport report = XDocReportRegistry.getRegistry().loadReport(is, TemplateEngineKind.Freemarker);
    	FieldsMetadata metadata = new FieldsMetadata();
    	report.setFieldsMetadata(metadata);
    	
    	IContext data = report.createContext();
    	data.put("sendername", "Antony Ho");
    	data.put("companyname", "Antony Workshop");
    	data.put("addrstreet", "11-13 Tai Yuen St");
    	data.put("addrpostalcode", "00000");
    	data.put("addrcity", "Kwai Chung");
    	data.put("addrstate", "HK");
    	
    	data.put("recipienttitle", "Lord");
    	data.put("recipientname", "Page Larry");
    	data.put("recipientstreet", "1600 Amphitheatre Parkway");
    	data.put("recipientpostalcode", "94043");
    	data.put("recipientcity", "Mountain View");
    	data.put("recipientstate", "CA");
    	
    	
    	/*
    	 * ***Method to add image into template***
    	 * in your docx and create a Bookmark linked to this image with the name "flag"
    	 * ("flag" is supposed to be the name of the image in this example)
    	 */
    	metadata.addFieldAsImage("flag");
    	IImageProvider flag = new FileImageProvider(new File(this.getClass().getResource("/flag2.png").getFile()));
    	data.put("flag", flag);
    	
    	// Output as DOCX
    	OutputStream docxOS = new FileOutputStream(new File("C:\\projects\\xdocreport-test\\testoutput.docx"));
    	report.process(data, docxOS);
    	
    	// Output as DOCX and then convert to PDF
    	OutputStream pdfOS = new FileOutputStream(new File("C:\\projects\\xdocreport-test\\testoutput.pdf"));
    	Options options = Options.getTo(ConverterTypeTo.PDF);
    	report.convert(data, options, pdfOS);
    	
    	
        assertTrue( true );
    }
}
