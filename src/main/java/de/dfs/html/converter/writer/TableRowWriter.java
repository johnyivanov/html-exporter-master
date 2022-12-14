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
package de.dfs.html.converter.writer;

import org.jsoup.nodes.Element;

public interface TableRowWriter {
	public static final String TABLE_ROW_ELEMENT_NAME = "tr";
	public static final String TABLE_BODY = "tbody";

	public static final String TD_TAG = "td";
	public static final String TH_TAG = "th";


	public void writeRow(Element element, int row);
}
