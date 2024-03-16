package cloud.huazai.documentprocessing.word;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * WordUtilsTest
 *
 * @author Di Wu
 * @since 2024-03-11
 */
public class WordUtilsTest {

	@Test
	public void demo() {


		XWPFDocument document = new XWPFDocument();

		String imgUrl = "https://mhp-test.oss-cn-qingdao.aliyuncs.com/images/test/ms/file/notFilename/20230920161635/20230920161634_3_.png";

		WordUtils.demo2(document, imgUrl);

		String outPath = "/Users/wudi/Downloads/HuaZai/测试3.docx";

		try {
			FileOutputStream fileOutputStream = new FileOutputStream(outPath);
			document.write(fileOutputStream);

			System.out.println("fileOutputStream = " + fileOutputStream);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}


	}
}