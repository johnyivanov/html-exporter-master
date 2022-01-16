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
package uk.co.certait.htmlexporter.demo;

import uk.co.certait.htmlexporter.writer.excel.ExcelExporter;

import java.io.File;
import java.nio.file.Files;
import java.util.stream.Collectors;

public class ReportGenerator {
	public ReportGenerator() throws Exception {
		File someFile = new File("report.html");
		String html;
		html = Files.lines(someFile.toPath()).collect(Collectors.joining(System.lineSeparator()));

		new ExcelExporter().exportHtml(html, new File("./report.xlsx"));
	}

	public static void main(String[] args) throws Exception {
		new ReportGenerator();

		System.exit(0);
	}

}
