/**
 * Copyright (C) 2012 alanhay <alanhay99@hotmail.com>
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *         http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package de.dfs.html.converter.writer.excel;

import de.dfs.html.converter.ss.CellRange;
import org.apache.poi.ss.usermodel.Cell;

import de.dfs.html.converter.ss.CellRangeObserver;
import de.dfs.html.converter.ss.CellRangeResolver;
import de.dfs.html.converter.ss.Function;

public class ExcelFunctionCell implements CellRangeObserver {
	private Cell cell;
	private CellRangeResolver resolver;
	private Function function;

	public ExcelFunctionCell(Cell cell, CellRange range, CellRangeResolver resolver, Function function) {
		this.cell = cell;
		this.resolver = resolver;
		this.function = function;

		range.addCellRangeObserver(this);

		cellRangeUpdated(range);
	}

	public void cellRangeUpdated(CellRange range) {
		cell.setCellFormula(function.toString() + "(" + resolver.getRangeString(range) + ")");
	}
}
