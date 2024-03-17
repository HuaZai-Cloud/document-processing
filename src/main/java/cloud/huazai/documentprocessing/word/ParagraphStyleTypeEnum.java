package cloud.huazai.documentprocessing.word;

import lombok.Getter;

/**
 * ParagraphStyleEnum
 *
 * @author Di Wu
 * @since 2024-03-17
 */
@Getter
public enum ParagraphStyleTypeEnum {


	MAIN_BODY("mainBody"),

	HEADING("Heading");


	private final String paragraphStyleName;

	ParagraphStyleTypeEnum(String paragraphStyleName) {
		this.paragraphStyleName = paragraphStyleName;
	}

}
