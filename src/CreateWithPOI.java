import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * Apache POI is an API that allows to create, modify and display Microsoft
 * Office files. This open source library is developed and distributed by Apache
 * Software Foundation. Apache POI releases are available under the Apache
 * License, Version 2.0.
 *
 * @author yuezhang
 *
 */
public class CreateWithPOI {
  public static void main(final String[] args) {
	final XMLSlideShow pptx = new XMLSlideShow();
	final File file = new File("d:/example.pptx");
	// Creating first slide
	// there can be multiple masters each referencing a number of
	// layouts for demonstration purposes we use the first (default)
	// slide master
	XSLFSlideMaster defaultMaster = pptx.getSlideMasters().get(0);
	XSLFSlideLayout slidelayout = defaultMaster.getLayout(SlideLayout.TITLE_ONLY);

	System.out.println("Adding picture");
	try {
	  // Draw picture as background
	  XSLFPictureData picture = pptx.addPicture(new File("d:/test.png"), PictureType.PNG);
	  XSLFPictureShape ps = slidelayout.createPicture(picture);
	  ps.setAnchor(new Rectangle2D.Double(100, 100, 400, 400));
	  pptx.createSlide(slidelayout);

	  // Draw picture
	  XSLFSlide slide = pptx.createSlide();
	  slide.createPicture(picture);
	} catch (IOException e1) {
	  // TODO Auto-generated catch block
	  e1.printStackTrace();
	}

	pptx.createSlide();

	pptx.getSlides().get(2).createTable();

	XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
	// fill the placeholders
	XSLFSlide slide1 = pptx.createSlide(titleLayout);
	XSLFTextShape title1 = slide1.getPlaceholder(0);
	title1.setText("First Title");
	try (FileOutputStream out = new FileOutputStream(file)) {
	  pptx.write(out);
	} catch (IOException e) {
	  System.out.println(e);
	}
  }
}
