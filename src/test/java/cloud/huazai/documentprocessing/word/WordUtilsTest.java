package cloud.huazai.documentprocessing.word;


import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
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
}