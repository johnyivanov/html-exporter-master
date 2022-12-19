/**
 * Copyright (C) 2012 alanhay <alanhay99@hotmail.com>
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package de.dfs.html.converter.writer.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import javax.swing.text.Style;
import java.util.HashMap;
import java.util.Map;


import static java.awt.Font.createFont;

public class ExcelStyleGenerator {
    private Map<Style, XSSFCellStyle> styles;

    public ExcelStyleGenerator() {
        styles = new HashMap<Style, XSSFCellStyle>();
    }


    public CellStyle getStyle(Cell cell, Style style) {
        XSSFCellStyle cellStyle;


        cellStyle = (XSSFCellStyle) cell.getSheet().getWorkbook().createCellStyle();


        //applyBackground(cell, cellStyle);
			/*applyBorders(style, cellStyle);
			applyFont(cell, style, cellStyle);
			applyHorizontalAlignment(style, cellStyle);
			applyverticalAlignment(style, cellStyle);
			applyWidth(cell, style);*/

        styles.put(style, (XSSFCellStyle) cellStyle);

        cellStyle.setWrapText(true);
        return cellStyle;
    }

    protected CellStyle getCellHeaderStyle(Workbook workbook, Cell cell, Sheet sheet) {
        CellStyle cellHeaderStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        font.setFontName("Courier New");
        font.setBold(true);
        cellHeaderStyle.setFont(font);

        cellHeaderStyle.setWrapText(true);
        cellHeaderStyle.setAlignment(HorizontalAlignment.CENTER);

        cellHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellHeaderStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());

        cellHeaderStyle.setBorderTop(BorderStyle.THIN);
        cellHeaderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellHeaderStyle.setBorderLeft(BorderStyle.THIN);
        cellHeaderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellHeaderStyle.setBorderBottom(BorderStyle.THIN);
        cellHeaderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellHeaderStyle.setBorderRight(BorderStyle.THIN);
        cellHeaderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

        return cellHeaderStyle;
    }

    protected CellStyle getCellStandartStyle(Workbook workbook, Cell cell, Sheet sheet) {
        CellStyle cellStandartStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Courier New");
        cellStandartStyle.setFont(font);
        cellStandartStyle.setWrapText(true);
        cellStandartStyle.setAlignment(HorizontalAlignment.CENTER);

        return cellStandartStyle;
    }


	/*protected CellStyle applyBackground(Cell cell, XSSFCellStyle cellStyle) {
		/*if (style.isBackgroundSet()) {
			cellStyle.setFillPattern(SOLID_FOREGROUND);
			cellStyle.setFillForegroundColor(new XSSFColor(style.getProperty(CssColorProperty.BACKGROUND)));
		}*/


/*
	protected void applyBorders(Style style, XSSFCellStyle cellStyle) {
		if (style.isBorderWidthSet()) {
			short width = (short) style.getProperty(CssIntegerProperty.BORDER_WIDTH);

			Color color = style.getProperty(CssColorProperty.BORDER_COLOR) != null ? style
					.getProperty(CssColorProperty.BORDER_COLOR) : Color.BLACK;

			cellStyle.setBorderBottom(BorderStyle.THIN);
			cellStyle.setBorderBottom(width);
			cellStyle.setBottomBorderColor(new XSSFColor(color));

			cellStyle.setBorderTop(BorderStyle.THIN);
			cellStyle.setBorderTop(width);
			cellStyle.setTopBorderColor(new XSSFColor(color));

			cellStyle.setBorderLeft(BorderStyle.THIN);
			cellStyle.setBorderLeft(width);
			cellStyle.setLeftBorderColor(new XSSFColor(color));

			cellStyle.setBorderRight(BorderStyle.THIN);
			cellStyle.setBorderRight(width);
			cellStyle.setRightBorderColor(new XSSFColor(color));
		}
	}

	protected void applyFont(Cell cell, Style style, XSSFCellStyle cellStyle) {
		Font font = createFont(cell.getSheet().getWorkbook(), style);
		cellStyle.setFont(font);
	}

	protected void applyHorizontalAlignment(Style style, XSSFCellStyle cellStyle) {
		if (style.isHorizontallyAlignedLeft()) {
			cellStyle.setAlignment(HorizontalAlignment.LEFT);
		} else if (style.isHorizontallyAlignedRight()) {
			cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		} else if (style.isHorizontallyAlignedCenter()) {
			cellStyle.setAlignment(HorizontalAlignment.CENTER);
		}
	}

	protected void applyverticalAlignment(Style style, XSSFCellStyle cellStyle) {
		if (style.isVerticallyAlignedTop()) {
			cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
		} else if (style.isVerticallyAlignedBottom()) {
			cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
		} else if (style.isVerticallyAlignedMiddle()) {
			cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		}
	}

	protected void applyWidth(Cell cell, Style style) {
		if (style.getProperty(CssIntegerProperty.WIDTH) > 0) {
			cell.getSheet().setColumnWidth(cell.getColumnIndex(), style.getProperty(CssIntegerProperty.WIDTH) * 50);
		}
	}

	public Font createFont(Workbook workbook, Style style) {
		Font font = workbook.createFont();

		if (style.isFontNameSet()) {
			font.setFontName(style.getProperty(CssStringProperty.FONT_FAMILY));
		}

		if (style.isFontSizeSet()) {
			font.setFontHeightInPoints((short) style.getProperty(CssIntegerProperty.FONT_SIZE));
		}

		if (style.isColorSet()) {
			Color color = style.getProperty(CssColorProperty.COLOR);

			// if(! color.equals(Color.WHITE)) // POI Bug
			// {
			((XSSFFont) font).setColor(new XSSFColor(color));
			// }
		}

		if (style.isFontBold()) {
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		}

		font.setItalic(style.isFontItalic());

		if (style.isTextUnderlined()) {
			font.setUnderline(Font.U_SINGLE);
		}

		return font;*/
}

