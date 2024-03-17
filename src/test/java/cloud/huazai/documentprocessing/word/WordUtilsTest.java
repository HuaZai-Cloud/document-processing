package cloud.huazai.documentprocessing.word;


import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

/**
 * WordUtilsTest
 *
 * @author Di Wu
 * @since 2024-03-11
 */
public class WordUtilsTest {

	@Test
	public void insertPictureTest() {


		XWPFDocument document = new XWPFDocument();

		String imgUrl = "https://mhp-test.oss-cn-qingdao.aliyuncs.com/images/test/ms/file/notFilename/20230920161635/20230920161634_3_.png";

		WordUtils.insertPicture(document, imgUrl,10,12,true, ParagraphAlignment.CENTER);

		String outPath = "/Users/wudi/Downloads/HuaZai/测试gl.docx";

		try {
			FileOutputStream fileOutputStream = new FileOutputStream(outPath);
			document.write(fileOutputStream);

		} catch (Exception e) {
			throw new RuntimeException(e);
		}


	}


	@Test
	public void insertPictureTest2() {


		XWPFDocument document = new XWPFDocument();

		String imgUrl = "https://mhp-test.oss-cn-qingdao.aliyuncs.com/images/test/ms/file/notFilename/20230920161635/20230920161634_3_.png";

		InputStream inputStream = null;
		try {
			URL url = new URL(imgUrl);
			URLConnection urlConn = url.openConnection();
			inputStream = urlConn.getInputStream();

			if (inputStream != null) {
				WordUtils.insertPicture(document, inputStream,14.6,12,ParagraphAlignment.CENTER);

				String outPath = "/Users/wudi/Downloads/HuaZai/测试cc.docx";
				FileOutputStream fileOutputStream = new FileOutputStream(outPath);
				document.write(fileOutputStream);
			}

		} catch (Exception e) {
			throw new RuntimeException(e);
		}


	}


	@Test
	void customStyle() {
		XWPFDocument document = new XWPFDocument();
		// WordUtils.customStyle(document,"22",1);


	}

	@Test
	void setHeading() {

		XWPFDocument document = new XWPFDocument();

		// WordUtils.setParagraphStyle(document,"标题1",1,30,"101010","宋体",2);
		// WordUtils.setParagraphStyle(document,"标题2",2,24,"F70E0E","仿宋",1.5);
		// WordUtils.setParagraphStyle(document,"标题3",3,12,"101010","微软雅黑",1.2);
		// WordUtils.setParagraphStyle(document,"正文",0,22,"101010","黑体",1.5);

		WordUtils.setText(document,"标题1",WordUtils.headingNameToStyleIdMap.get("标题1"));
		WordUtils.setText(document,"一个文档可以有多个页眉, 页眉里面可以包含段落和表格,获取文档的页眉：List headerList = doc.getHeaderList();获取页眉里的所有段落：List paras = header.getParagraphs();获取页眉里的所有表格：List tables = header.getTables();",WordUtils.headingNameToStyleIdMap.get("正文"));
		WordUtils.setText(document,"标题2",WordUtils.headingNameToStyleIdMap.get("标题2"));
		WordUtils.setText(document,"页脚和页眉基本类似，可以获取表示页数的角标",WordUtils.headingNameToStyleIdMap.get("正文"));
		WordUtils.setText(document,"标题3",WordUtils.headingNameToStyleIdMap.get("标题3"));
		WordUtils.setText(document,"直接调用 XWPFRun 的 setText() 方法设置文本时，在底层会重新创建一个 XWPFRun，把文本附加在当前文本后面，所以我们不能直接设值，需要先删除当前 run, 然后再自己手动插入一个新的 run",WordUtils.headingNameToStyleIdMap.get("正文"));


		String outPath = "/Users/wudi/Downloads/HuaZai/测试设置标题格式5.docx";

		try {
			FileOutputStream fileOutputStream = new FileOutputStream(outPath);
			document.write(fileOutputStream);

		} catch (Exception e) {
			throw new RuntimeException(e);
		}


	}
}