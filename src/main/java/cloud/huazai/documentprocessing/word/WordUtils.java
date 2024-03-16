package cloud.huazai.documentprocessing.word;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

/**
 * WordUtils
 *
 * @author Di Wu
 * @since 2024-03-11
 */
public class WordUtils {

	// public static void  download(XWPFDocument document, HttpServletResponse response, String wordName)throws Exception{
	// 	String attachmentPath = wordName + ".docx";
	// 	String fileNameURL = URLEncoder.encode(attachmentPath, "UTF-8");
	// 	response.setContentType("application/octet-stream;charset=UTF-8");
	// 	response.setHeader("Content-disposition", "attachment;filename=" + fileNameURL + ";" + "filename*=" +fileNameURL);
	// 	response.setCharacterEncoding("UTF-8");
	//
	// 	document.write(response.getOutputStream());
	// }


	public static void demo(XWPFDocument document,String imgUrl){

		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();

		int format;
		if (imgUrl.endsWith(".emf")) {
			format = XWPFDocument.PICTURE_TYPE_EMF;
		} else if (imgUrl.endsWith(".wmf")) {
			format = XWPFDocument.PICTURE_TYPE_WMF;
		} else if (imgUrl.endsWith(".pict")) {
			format = XWPFDocument.PICTURE_TYPE_PICT;
		} else if (imgUrl.endsWith(".jpeg") || imgUrl.endsWith(".jpg")) {
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		} else if (imgUrl.endsWith(".png")) {
			format = XWPFDocument.PICTURE_TYPE_PNG;
		} else if (imgUrl.endsWith(".dib")) {
			format = XWPFDocument.PICTURE_TYPE_DIB;
		} else if (imgUrl.endsWith(".gif")) {
			format = XWPFDocument.PICTURE_TYPE_GIF;
		} else if (imgUrl.endsWith(".tiff")) {
			format = XWPFDocument.PICTURE_TYPE_TIFF;
		} else if (imgUrl.endsWith(".eps")) {
			format = XWPFDocument.PICTURE_TYPE_EPS;
		} else if (imgUrl.endsWith(".bmp")) {
			format = XWPFDocument.PICTURE_TYPE_BMP;
		} else if (imgUrl.endsWith(".wpg")) {
			format = XWPFDocument.PICTURE_TYPE_WPG;
		} else {
			return;
		}
		int index=0;
		try {
			URL	url = new URL(imgUrl);
			URLConnection urlConn = url.openConnection();
			InputStream inputStream = urlConn.getInputStream();


			if (null != inputStream) {
				String relationshipId = paragraph.getDocument().getPackagePart()
						.addExternalRelationship(imgUrl, XWPFRelation.HYPERLINK.getRelation()).getId();

				run.addPicture(inputStream, format,"图片", Units.toEMU(440), Units.toEMU(230));

				if (run.getCTR().getDrawingList() != null && !run.getCTR().getDrawingList().isEmpty()) {
					CTDrawing ctDrawing = run.getCTR().getDrawingList().get(index);
					if (ctDrawing.getInlineList() != null && !ctDrawing.getInlineList().isEmpty()) {
						CTInline ctInline = ctDrawing.getInlineList().get(0);
						CTNonVisualDrawingProps docPr = ctInline.getDocPr();
						if (docPr != null) {
							org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink hlinkClick = docPr.addNewHlinkClick();
							hlinkClick.setId(relationshipId);
						}
					}
					index++;
				}
			}
		} catch (Exception e) {
			throw new RuntimeException(e);
		}




	}

	public static void demo2(XWPFDocument document,String imgUrl){


		int format;
		if (imgUrl.endsWith(".emf")) {
			format = XWPFDocument.PICTURE_TYPE_EMF;
		} else if (imgUrl.endsWith(".wmf")) {
			format = XWPFDocument.PICTURE_TYPE_WMF;
		} else if (imgUrl.endsWith(".pict")) {
			format = XWPFDocument.PICTURE_TYPE_PICT;
		} else if (imgUrl.endsWith(".jpeg") || imgUrl.endsWith(".jpg")) {
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		} else if (imgUrl.endsWith(".png")) {
			format = XWPFDocument.PICTURE_TYPE_PNG;
		} else if (imgUrl.endsWith(".dib")) {
			format = XWPFDocument.PICTURE_TYPE_DIB;
		} else if (imgUrl.endsWith(".gif")) {
			format = XWPFDocument.PICTURE_TYPE_GIF;
		} else if (imgUrl.endsWith(".tiff")) {
			format = XWPFDocument.PICTURE_TYPE_TIFF;
		} else if (imgUrl.endsWith(".eps")) {
			format = XWPFDocument.PICTURE_TYPE_EPS;
		} else if (imgUrl.endsWith(".bmp")) {
			format = XWPFDocument.PICTURE_TYPE_BMP;
		} else if (imgUrl.endsWith(".wpg")) {
			format = XWPFDocument.PICTURE_TYPE_WPG;
		} else {
			return;
		}
		int index=0;
		try {
			URL	url = new URL(imgUrl);
			URLConnection urlConn = url.openConnection();
			InputStream inputStream = urlConn.getInputStream();
			String id = document.addPictureData(inputStream, format);

			XWPFParagraph paragraph = document.createParagraph();
			XWPFRun run = paragraph.createRun();


			if (null != inputStream) {

				// String relationshipId = paragraph.getDocument().getPackagePart()
				// 		.addExternalRelationship(imgUrl, XWPFRelation.HYPERLINK.getRelation()).getId();

				run.addPicture(inputStream, format,"图片", Units.toEMU(440), Units.toEMU(230));

				if (run.getCTR().getDrawingList() != null && !run.getCTR().getDrawingList().isEmpty()) {
					CTDrawing ctDrawing = run.getCTR().getDrawingList().get(index);
					if (ctDrawing.getInlineList() != null && !ctDrawing.getInlineList().isEmpty()) {
						CTInline ctInline = ctDrawing.getInlineList().get(0);
						CTNonVisualDrawingProps docPr = ctInline.getDocPr();
						if (docPr != null) {
							org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink hlinkClick = docPr.addNewHlinkClick();
							hlinkClick.setId(id);
						}
					}
					index++;
				}
			}
		} catch (Exception e) {
			throw new RuntimeException(e);
		}


	}


}
