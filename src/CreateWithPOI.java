import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPlaceholder;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;
import org.openxmlformats.schemas.presentationml.x2006.main.STPlaceholderType;

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
	createWithDefaultTemplate();
	createWithCustomisedTemplate();
	// useTemplateToCreateNewPPT();
  }

  private static void createWithDefaultTemplate() {

	final XMLSlideShow pptx = new XMLSlideShow();
	// Creating first slide
	// there can be multiple masters each referencing a number of
	// layouts for demonstration purposes we use the first (default)
	// slide master
	XSLFSlideMaster defaultMaster = pptx.getSlideMasters().get(0);
	XSLFSlideLayout slidelayout = defaultMaster.getLayout(SlideLayout.TITLE_ONLY);

	try {
	  // Draw picture as background
	  XSLFPictureData picture = pptx.addPicture(new File("d:/test.png"), PictureType.PNG);
	  XSLFPictureShape ps = slidelayout.createPicture(picture);
	  ps.setAnchor(new Rectangle2D.Double(100, 100, 400, 400));
	  pptx.createSlide(slidelayout);

	  // Insert picture
	  XSLFSlide slide = pptx.createSlide(slidelayout);
	  XSLFTextShape title = slide.getPlaceholder(0);
	  title.setText("picture title");
	  XSLFPictureShape shape = slide.createPicture(picture);
	  shape.setAnchor(new Rectangle(100, 100, 400, 400));

	  // Insert picture and scale respect ratio
	  XSLFSlide slide2 = pptx.createSlide();
	  XSLFPictureShape shape2 = slide2.createPicture(picture);
	  final Dimension availableSize = pptx.getPageSize();
	  final Dimension scaledSize = getScaledDimension((float) picture.getImageDimension().getWidth(), (float) picture
		  .getImageDimension().getHeight(), availableSize, 0.8f);
	  final int x = (availableSize.width - scaledSize.width) / 2;
	  final int y = (availableSize.height - scaledSize.height) / 2;
	  shape2.setAnchor(new Rectangle(x, y, scaledSize.width, scaledSize.height));
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

	final File file = new File("d:/example.pptx");
	try (FileOutputStream out = new FileOutputStream(file)) {
	  pptx.write(out);
	} catch (IOException e) {
	  System.out.println(e);
	}
  }

  private static void createWithCustomisedTemplate() {
	try (final FileInputStream templateStream = new FileInputStream(new File("d:/test.potx"))) {
	  final XMLSlideShow pptx = new XMLSlideShow(templateStream);

	  final XSLFSlideMaster defaultMaster = pptx.getSlideMasters().get(0);
	  final XSLFSlideLayout templateLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

	  XSLFPictureData picture = pptx.addPicture(new File("d:/test.png"), PictureType.PNG);

	  XSLFSlide slide = pptx.createSlide(templateLayout);
	  XSLFTextShape title = slide.getPlaceholder(0);
	  title.setText("picture title");

	  // Get anchor of the content place holder
	  Rectangle2D anchor = slide.getPlaceholder(1).getAnchor();
	  XSLFPictureShape shape = slide.createPicture(picture);
	  // Remove default shape
	  slide.removeShape(slide.getPlaceholder(1));
	  // Replace it with the picture
	  shape.setAnchor(anchor);

	  pptx.getPackage().replaceContentType(XSLFRelation.PRESENTATIONML_TEMPLATE.getContentType(),
		  XSLFRelation.MAIN.getContentType());

	  try (FileOutputStream out = new FileOutputStream(new File("d:/example2.pptx"))) {
		pptx.write(out);
	  } catch (IOException e) {
		System.out.println(e);
	  }
	} catch (FileNotFoundException e) {
	  // TODO Auto-generated catch block
	  e.printStackTrace();
	} catch (IOException e1) {
	  // TODO Auto-generated catch block
	  e1.printStackTrace();
	}
  }

  private static void useTemplateToCreateNewPPT() {
	try (final FileInputStream templateStream = new FileInputStream(new File("d:/test.potx"))) {
	  final XMLSlideShow potx = new XMLSlideShow(templateStream);
	  final XSLFSlideMaster defaultMaster = potx.getSlideMasters().get(0);
	  final XSLFSlideLayout templateLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

	  XMLSlideShow pptx = new XMLSlideShow();
	  XSLFPictureData picture = pptx.addPicture(new File("d:/test.png"), PictureType.PNG);
	  XSLFSlide slide = pptx.createSlide(templateLayout);

	  XSLFTextShape title = slide.getPlaceholder(0);
	  title.setText("picture title");
	  XSLFPictureShape shape = slide.createPicture(picture);
	  shape.setAnchor(new Rectangle(100, 100, 400, 400));

	  try (FileOutputStream out = new FileOutputStream(new File("d:/example3.pptx"))) {
		pptx.write(out);
	  } catch (IOException e) {
		System.out.println(e);
	  }
	} catch (FileNotFoundException e) {
	  // TODO Auto-generated catch block
	  e.printStackTrace();
	} catch (IOException e1) {
	  // TODO Auto-generated catch block
	  e1.printStackTrace();
	}
  }

  private static XSLFShape findPicPlaceHolder(final XSLFSlide slide) {
	List<XSLFShape> shapes = slide.getShapes();
	for (XSLFShape shape : shapes) {

	  System.out.println(shape.getShapeName());

	  if (shape instanceof XSLFTextShape) {
		System.out.println("TEXTSHAPE");

		CTShape sh = (CTShape) shape.getXmlObject();
		CTPlaceholder ph = sh.getNvSpPr().getNvPr().getPh();

		if (ph != null) {
		  if (ph.getType() == STPlaceholderType.PIC) {
			return shape;
		  }
		}
	  }
	}
	return null;
  }

  /**
   * Resize to boundary while maintain aspect ratio
   *
   * @param imageWidth
   * @param imageHeight
   * @param boundary
   * @return
   */
  private static Dimension getScaledDimension(final float imageWidth, final float imageHeight,
	  final Dimension boundary, final float margin) {
	final float bound_width = boundary.width * margin;
	final float bound_height = boundary.height * margin;
	Dimension newDim = new Dimension();
	double scaleX = bound_width / imageWidth;
	double scaleY = bound_height / imageHeight;
	double scale = Math.min(scaleX, scaleY);
	newDim.width = (int) Math.round(imageWidth * scale);
	newDim.height = (int) Math.round(imageHeight * scale);

	return newDim;
  }
}
