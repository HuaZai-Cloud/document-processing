package cloud.huazai.documentprocessing.word;

import lombok.Getter;

/**
 * LineSpacingEnum
 *
 * @author Di Wu
 * @since 2024-03-17
 */
@Getter
public enum LineSpacingEnum {

	one(1),
	one_point_five(1.5),

	two(2),
	two_point_five(2.5),

	three(3),

	three_point_five(3.5),

	four(4),

	four_point_five(4.5),

	five(5),
	;

	private final double lineSpacing;

	LineSpacingEnum(double lineSpacing) {
		this.lineSpacing = lineSpacing;
	}

}
