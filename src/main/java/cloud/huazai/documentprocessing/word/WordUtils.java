package cloud.huazai.documentprocessing.word;

import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

import java.io.BufferedInputStream;
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

	/**
	 * 插入图片
	 *
	 * @param document 文档
	 * @param pictureUrl 图片url
	 * @param width 文档中图片宽度 cm
	 * @param height 文档中图片高度 cm
	 * @param isAddConnection 是否添加超链接
	 * @param alignment 对齐方式
	 */
	public static void insertPicture(XWPFDocument document,String pictureUrl,double width, double height ,boolean isAddConnection,ParagraphAlignment alignment){
		try {
			URL	url = new URL(pictureUrl);
			URLConnection urlConn = url.openConnection();
			InputStream inputStream = urlConn.getInputStream();

			XWPFParagraph paragraph = document.createParagraph();
			XWPFRun run = paragraph.createRun();

			if (null != inputStream) {

				BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);

				PictureType pictureType = PictureType.valueOf(FileMagic.valueOf(bufferedInputStream));
				int format = pictureType.getOoxmlId();

				paragraph.setAlignment(alignment);

				run.addPicture(bufferedInputStream, format,"picture", (int) Math.rint(width * Units.EMU_PER_CENTIMETER),(int)Math.rint(height*Units.EMU_PER_CENTIMETER));

				if (isAddConnection) {

					String relationshipId = document.getPackagePart().addExternalRelationship(pictureUrl, XWPFRelation.HYPERLINK.getRelation()).getId();

					if (run.getCTR().getDrawingList() != null && !run.getCTR().getDrawingList().isEmpty()) {
						CTDrawing ctDrawing = run.getCTR().getDrawingList().get(0);
						if (ctDrawing.getInlineList() != null && !ctDrawing.getInlineList().isEmpty()) {
							CTInline ctInline = ctDrawing.getInlineList().get(0);
							CTNonVisualDrawingProps docPr = ctInline.getDocPr();
							if (docPr != null) {
								CTHyperlink linkClick = docPr.addNewHlinkClick();
								linkClick.setId(relationshipId);
							}
						}
					}
				}
			}
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}


	/**
	 * 插入图片
	 *
	 * @param document 文档
	 * @param pictureInputStream 图片输入流
	 * @param width 文档中图片宽度 cm
	 * @param height 文档中图片高度 cm
	 * @param alignment 对齐方式
	 */
	public static void insertPicture(XWPFDocument document,InputStream pictureInputStream,double width, double height ,ParagraphAlignment alignment){

		try {

			XWPFParagraph paragraph = document.createParagraph();
			XWPFRun run = paragraph.createRun();

			if (null != pictureInputStream) {

				BufferedInputStream bufferedInputStream = new BufferedInputStream(pictureInputStream);

				PictureType pictureType = PictureType.valueOf(FileMagic.valueOf(bufferedInputStream));
				int format = pictureType.getOoxmlId();
				paragraph.setAlignment(alignment);

				run.addPicture(bufferedInputStream, format,"picture", (int) Math.rint(width * Units.EMU_PER_CENTIMETER),(int)Math.rint(height*Units.EMU_PER_CENTIMETER));
			}
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}



}
