package edu.handong.csee.isel;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.extractor.POIOLE2TextExtractor;
import org.apache.poi.extractor.POITextExtractor;
import org.apache.poi.ooxml.extractor.ExtractorFactory;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.xmlbeans.XmlException;

public class Main {
	static String path = "/Users/imseongbin/Desktop/ppt/in/";
	static String outPath = "/Users/imseongbin/Desktop/ppt/out/";
	static XMLSlideShow samplePPT;// = new XMLSlideShow(new FileInputStream(new File("/Users/imseongbin/Desktop/ppt/sample.pptx")));
	
	static XSLFSlide getThreeLineSlide() {
		return samplePPT.getSlides().get(0);
	}
	static XSLFSlide getTwoLineSlide() {
		return samplePPT.getSlides().get(1);
	}
	
	public static void main(String[] args) throws IOException, OpenXML4JException, XmlException {
		File dirIn = new File(path);
		init();
		
		for(File file : dirIn.listFiles()) {
			String name = file.getName();
			if(!name.contains("pptx"))
				continue;
			System.out.println(file.getAbsolutePath());
			System.out.println(name);
			FileInputStream in = new FileInputStream(file);
			XMLSlideShow ppt = new XMLSlideShow(in);
			File outFile = new File(outPath + name);
			if(outFile.exists()) outFile.delete();
			XMLSlideShow newPPT = new XMLSlideShow();
			List<XSLFSlide> slides = ppt.getSlides();
			
			for(XSLFSlide slide : slides) {
				XSLFTextShape txtsh = (XSLFTextShape) slide.getShapes().get(0);
				String text = txtsh.getText();
				String[] lines = text.split("\n");
				int countLine = lines.length;
				XSLFSlide newSlide = null;
				if(countLine == 2) {
					newSlide = newPPT.createSlide().importContent(getTwoLineSlide());
				} else if(countLine >= 3) {
					newSlide = newPPT.createSlide().importContent(getThreeLineSlide());
				}
				XSLFTextShape newtxtsh = (XSLFTextShape) newSlide.getShapes().get(0);
				XSLFTextRun run = newtxtsh.setText(lines[0]);
				run.setBold(true);
				run.setFontSize(44.0);
				for(int i = 1; i < lines.length ; i++) {
					TextParagraph para = newtxtsh.addNewTextParagraph();
					XSLFTextRun newrun = newtxtsh.appendText(lines[i], false);
					newrun.setBold(true);
					newrun.setFontSize(44.0);
					para.setTextAlign(TextAlign.CENTER);
				}
				
				
			}
			newPPT.write(new FileOutputStream(outFile));
			
			
			
		}
		
	}

	private static void init() throws FileNotFoundException, IOException {
		samplePPT = new XMLSlideShow(new FileInputStream(new File("/Users/imseongbin/Desktop/ppt/sample.pptx")));
		
	}

}
