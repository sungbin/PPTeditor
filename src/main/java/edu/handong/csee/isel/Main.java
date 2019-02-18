package edu.handong.csee.isel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.xmlbeans.XmlException;

public class Main {
	static String path = "/Users/imseongbin/Desktop/ppt/in/";
	static String outPath = "/Users/imseongbin/Desktop/ppt/out/";
	static String samplePath = "/Users/imseongbin/Desktop/ppt/sample.pptx";
	static XMLSlideShow samplePPT;// = new XMLSlideShow(new FileInputStream(new
									// File("/Users/imseongbin/Desktop/ppt/sample.pptx")));

	static XSLFSlide getThreeLineSlide() {
		return samplePPT.getSlides().get(0);
	}

	static XSLFSlide getTwoLineSlide() {
		return samplePPT.getSlides().get(1);
	}

	public static void main(String[] args) throws IOException, OpenXML4JException, XmlException {
		File dirIn = new File(path);
		init();

		for (File file : dirIn.listFiles()) {
			String name = file.getName();
			if (!name.contains("pptx"))
				continue;
//			System.out.println(name);
			FileInputStream in = new FileInputStream(file);
			XMLSlideShow ppt = new XMLSlideShow(in);
			File outFile = new File(outPath + name);
			if (outFile.exists())
				outFile.delete();
			XMLSlideShow newPPT = new XMLSlideShow();
			List<XSLFSlide> slides = ppt.getSlides();
			int slideCount = 0;
			for (XSLFSlide slide : slides) {
				List<XSLFShape> txtshlst = slide.getShapes();
				if (txtshlst.size() > 1)
					System.out.println("Warning: Many Text box Slide " + (slideCount + 1) + "st in " + file.getAbsolutePath());
				StringBuffer sb = new StringBuffer();
				for (XSLFShape sh : txtshlst) {
					XSLFTextShape txtsh = (XSLFTextShape) sh;
					String text = txtsh.getText();
					sb.append(text + "\n");
				}

				String text = sb.toString();
				String[] lines = text.split("\n");
				int countLine = lines.length;
				XSLFSlide newSlide = null;
				if (countLine == 2) {
					newSlide = newPPT.createSlide().importContent(getTwoLineSlide());
				} else if (countLine >= 3) {
					newSlide = newPPT.createSlide().importContent(getThreeLineSlide());
				}
				XSLFTextShape newtxtsh = (XSLFTextShape) newSlide.getShapes().get(0);
				XSLFTextRun run = newtxtsh.setText(lines[0]);
				run.setBold(true);
				run.setFontSize(48.0);
				run.setFontFamily("HY엽서M");
				run.setCharacterSpacing(-3.0);
				for (int i = 1; i < lines.length; i++) {
//					TextParagraph para = newtxtsh.addNewTextParagraph();
					XSLFTextRun newrun = newtxtsh.appendText("\n"+lines[i], false);
//					para.setBulletStyle("HY엽서M");
//					para.setTextAlign(TextAlign.CENTER);
					newrun.setBold(true);
					newrun.setFontSize(48.0);
					newrun.setCharacterSpacing(-3.0);
					newrun.setFontFamily("HY엽서M");
				}

			}
			newPPT.write(new FileOutputStream(outFile));
		}
	}

	private static void init() throws FileNotFoundException, IOException {
		samplePPT = new XMLSlideShow(new FileInputStream(new File(samplePath)));
	}
}
