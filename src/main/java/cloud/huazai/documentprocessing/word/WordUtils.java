package cloud.huazai.documentprocessing.word;

import cloud.huazai.tool.basic.lang.StringUtils;
import jdk.internal.org.objectweb.asm.Handle;
import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.xml.bind.annotation.adapters.HexBinaryAdapter;
import java.io.BufferedInputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URL;
import java.net.URLConnection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * WordUtils
 *
 * @author Di Wu
 * @since 2024-03-11
 */
public class WordUtils {
	private static final String pictureName = "picture";

	public static Map<String, String> headingNameToStyleIdMap = new HashMap<>();

	public static void insertPicture(XWPFDocument document, String pictureUrl, double width, double height) {
		insertPicture(document, pictureUrl, width, height, true, ParagraphAlignment.CENTER);
	}

	public static void insertPicture(XWPFDocument document, InputStream pictureInputStream, double width, double height) {
		insertPicture(document, pictureInputStream, width, height, ParagraphAlignment.CENTER);
	}


	/**
	 * 插入图片
	 *
	 * @param document        文档
	 * @param pictureUrl      图片url
	 * @param width           文档中图片宽度 cm
	 * @param height          文档中图片高度 cm
	 * @param isAddConnection 是否添加超链接
	 * @param alignment       对齐方式
	 */
	public static void insertPicture(XWPFDocument document, String pictureUrl, double width, double height, boolean isAddConnection, ParagraphAlignment alignment) {
		try {
			URL url = new URL(pictureUrl);
			URLConnection urlConn = url.openConnection();
			InputStream inputStream = urlConn.getInputStream();
			XWPFParagraph paragraph = document.createParagraph();
			XWPFRun run = paragraph.createRun();
			if (inputStream != null) {
				BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);

				PictureType pictureType = PictureType.valueOf(FileMagic.valueOf(bufferedInputStream));
				int format = pictureType.getOoxmlId();

				paragraph.setAlignment(alignment);

				run.addPicture(bufferedInputStream, format, pictureName, (int) Math.rint(width * Units.EMU_PER_CENTIMETER), (int) Math.rint(height * Units.EMU_PER_CENTIMETER));

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
	 * @param document           文档
	 * @param pictureInputStream 图片输入流
	 * @param width              文档中图片宽度 cm
	 * @param height             文档中图片高度 cm
	 * @param alignment          对齐方式
	 */
	public static void insertPicture(XWPFDocument document, InputStream pictureInputStream, double width, double height, ParagraphAlignment alignment) {
		try {
			XWPFParagraph paragraph = document.createParagraph();
			XWPFRun run = paragraph.createRun();
			if (null != pictureInputStream) {
				BufferedInputStream bufferedInputStream = new BufferedInputStream(pictureInputStream);
				PictureType pictureType = PictureType.valueOf(FileMagic.valueOf(bufferedInputStream));
				int format = pictureType.getOoxmlId();
				paragraph.setAlignment(alignment);
				run.addPicture(bufferedInputStream, format, pictureName, (int) Math.rint(width * Units.EMU_PER_CENTIMETER), (int) Math.rint(height * Units.EMU_PER_CENTIMETER));
			}
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}


	public static void setHeading(XWPFDocument document,String headingName,int headingLevel,int fontSize,String fontColor, String fontName,double lineSpacing) {

		XWPFStyles styles = document.createStyles();

		//创建样式
		CTStyle ctStyle = CTStyle.Factory.newInstance();


		//设置id
		ctStyle.setStyleId("Heading"+headingLevel);
		// ctStyle.setStyleId(headingName);

		headingNameToStyleIdMap.put(headingName, ctStyle.getStyleId());


		CTString styleName = CTString.Factory.newInstance();
		styleName.setVal(ctStyle.getStyleId());
		ctStyle.setName(styleName);

		CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
		indentNumber.setVal(BigInteger.valueOf(headingLevel));

		// 数字越低在格式栏中越突出
		ctStyle.setUiPriority(indentNumber);

		CTOnOff onoffnull = CTOnOff.Factory.newInstance();
		ctStyle.setUnhideWhenUsed(onoffnull);

		// 样式将显示在“格式”栏中
		ctStyle.setQFormat(onoffnull);

		// 样式定义给定级别的标题
		if (headingLevel != 0) {
			CTPPrGeneral ppr = CTPPrGeneral.Factory.newInstance();
			ppr.setOutlineLvl(indentNumber);
			ctStyle.setPPr(ppr);
		}

		setSingleLineSpacing(ctStyle, lineSpacing);


		XWPFStyle style = new XWPFStyle(ctStyle);

		CTRPr rpr = CTRPr.Factory.newInstance();


		CTHpsMeasure sz = rpr.addNewSz();
		sz.setVal(new BigInteger(String.valueOf(fontSize)));

		CTHpsMeasure szCs = rpr.addNewSzCs();
		szCs.setVal(new BigInteger(String.valueOf(fontSize)));


		CTFonts fonts = rpr.addNewRFonts();
		if (StringUtils.isNotBlank(fontName)) {
			fontName = "微软雅黑";
		}
		fonts.setAscii(fontName);

		CTColor color = rpr.addNewColor();

		color.setVal(hexToBytes(fontColor));
		style.getCTStyle().setRPr(rpr);




		style.setType(STStyleType.PARAGRAPH);




		styles.addStyle(style);
	}

	public static byte[] hexToBytes(String hexString) {
		HexBinaryAdapter adapter = new HexBinaryAdapter();
		return adapter.unmarshal(hexString);
	}



	// int styleLevel
	public static void setHeading(XWPFDocument document, String text, String style, String color, Integer fontAlignment,
	                           Integer fontSize, Boolean bold, Integer changeLine, Boolean changePage,
	                           int indentationFirstLine,double lineSpacing) {
		XWPFParagraph paragraph = document.createParagraph();
		paragraph.setIndentationFirstLine(indentationFirstLine);
		//
		paragraph.setFontAlignment(fontAlignment);
		//标题样式
		if (StringUtils.isNotBlank(style)) {
			XWPFRun run = paragraph.createRun();
			run.setText(text);
			run.setFontSize(fontSize);
			run.setBold(bold);
			run.setFontFamily("宋体");
			if (StringUtils.isNotBlank(color)) {
				run.setColor(color);
			}
			setSingleLineSpacing(paragraph,lineSpacing);
			paragraph.setStyle("Heading" + style); // 标题 2
			// paragraph.setStyle("标题 " + style); // 标题 2
			return;
		}

		XWPFRun run = paragraph.createRun();
		if (StringUtils.isNotBlank(color)) {
			run.setColor(color);
		}
		//字体
		run.setFontSize(fontSize);
		run.setBold(bold);
		run.setText(text);
		run.setFontFamily("宋体");
		setSingleLineSpacing(paragraph,lineSpacing);
		//换行
		for (int i = 0; i < changeLine; i++) {
			run.addCarriageReturn();
		}
		//换页
		if (changePage) {
			run.addBreak(BreakType.PAGE);
		}
	}



	public static void setSingleLineSpacing(XWPFParagraph para, double lineSpacing) {

		if (lineSpacing <= 0) {
			lineSpacing = 1;
		}

		CTPPr ppr = para.getCTP().getPPr();
		if (ppr == null) ppr = para.getCTP().addNewPPr();
		CTSpacing spacing = ppr.isSetSpacing()? ppr.getSpacing() : ppr.addNewSpacing();
		spacing.setAfter(BigInteger.valueOf(0));
		spacing.setBefore(BigInteger.valueOf(0));
		spacing.setLineRule(STLineSpacingRule.AUTO);
		int line = (int) (24 * lineSpacing * 10);
		spacing.setLine(BigInteger.valueOf(line));
	}

	public static void setSingleLineSpacing(CTStyle ctStyle, double lineSpacing) {

		if (lineSpacing <= 0) {
			lineSpacing = 1;
		}

		CTPPrGeneral ppr = ctStyle.getPPr();
		if (ppr == null) {ppr = ctStyle.addNewPPr();}
		CTSpacing spacing = ppr.isSetSpacing()? ppr.getSpacing() : ppr.addNewSpacing();
		spacing.setAfter(BigInteger.valueOf(0));
		spacing.setBefore(BigInteger.valueOf(0));
		spacing.setLineRule(STLineSpacingRule.AUTO);
		int line = (int) (24 * lineSpacing * 10);
		spacing.setLine(BigInteger.valueOf(line));
	}


	public static void setText(XWPFDocument document, String text,
	                            String styleId) {
		XWPFParagraph paragraph = document.createParagraph();
		// paragraph.setIndentationFirstLine(indentationFirstLine);
		//
		// paragraph.setFontAlignment(fontAlignment);

		paragraph.setStyle(styleId);
		XWPFRun run = paragraph.createRun();
		// run.setStyle(styleId);
		// if (StringUtils.isNotBlank(color)) {
		// 	run.setColor(color);
		// }
		//字体
		// run.setFontSize(fontSize);
		// run.setBold(bold);
		run.setText(text);
		// run.setFontFamily("宋体");
		// setSingleLineSpacing(paragraph,lineSpacing);

	}

}
