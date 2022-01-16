html-exporter 
============

Java Library to Export CSS styled HTML to various Formats (XLSX, ODS, PDF, DOCX etc)

Sample Reports generated using the library:

<a href ="http://tinyurl.com/nhq5mu9">Open Office ODS Spreadsheet</a><br/>
<a href ="http://tinyurl.com/pbnao9u">MS Excel XLSX Spreadsheet</a><br/>
<a href ="http://tinyurl.com/o2hk9l7">PDF Document</a>


Initial development has concentrated around  Excel (via POI) and Open Office Calc (via the ODS Toolkit)
as I couldn't find an existing HTML to (proper) spreadsheet (as opposed to csv) exporter for java.

For PDF generation this library simply provides some convenience methods around the existing 
<a href="https://code.google.com/p/flying-saucer/">xhtmlRenderer/flying saucer</a> library. 

Word and OO Writer both open HTML documents regardless of the file extension so whether this is worth looking at remains to be seen.

Usage
=====

The main purpose of the library is for use in reporting. While there are existing Java reporting frameworks 
(e.g. Jasper and BIRT) these have a fairly steep learning curve and can be painful to use: 

- lay out every report using the IDE tools
- use some cryptic scripting language to control conditional behaviour
- create a bunch of DTOs to provide the data.

If you are not using the report server (and possibly viewer) features of these reporting frameworks then 
the main feature they offer is the convenience of write once export to multiple formats.

An alternative approach then is to generate the report in HTML/CSS using you favourite easy to use templating library (Velocity, Freemarker, 
StringTemplate etc.) and use this library to export to the various formats which end-users would expect.

Example usage:

	String html = generateMyReport();

	new ExcelExporter().exportHtml(html, new File("./report.xlsx"));
	//or byte [] = new ExcelExporter().exportHtml(html);
	
	new OdsExporter().exportHtml(html, new File("./report.ods"));
	//or byte [] = new ExcelExporter().exportHtml(html);
	
	new PdfExporter().exportHtml(html, new File("./report.pdf"));
	//or byte [] = new ExcelExporter().exportHtml(html);
	

The reports linked to above were generated via a demo application which is included in the source. This generates some data used 
to populate a Velocity template. The resulting HTML is then exported to Excel, PDF and Open Office Calc.

Features
========

* Supports a limited subset of CSS: enough for producing nicely formatted reports.
* Spreadsheet cell comments.
* Cells can span multiple rows and columns.
* Multiple sheets.
* Automatic formula insertion:
* Freeze panes

The ODS and Excel exporters allow for producing spreadsheets with automatic formula insertion via the use of various HTML5 
compliant data-* attributes being applied to table cells. The sample spreadsheets linked to above demonstrate this functionality.

Example:

	<!-- This cell will, via the data-group attribute,  be added to two ranges, each of which will be the inputs to formulas -->
	<td data-group="store_Dumfries_2_value, region_1_1_pg_5_value" class="numeric">
    	486
    </td>      
    
    ...
	
	<!-- 
		This raw value of this cell will, via the data-group-output attribute, be replaced with a SUM function taking as input all cells added to the specified range.
		The cell is then itself added to another range which will be used by a further function.
	-->
    <td class="subTotal numeric" data-group-output="region_1_1_pg_6_count" data-group="area_1_pg_6_count">
		32
    </td>
    
Maven
=====

To try out the library quickly add the following repository to your POM:

		<repositories>
			<repository>
				<id>maven-s3-release-repo</id>
				<name>S3 Release Repository</name>
				<url>s3://alanhay-maven-repository/release</url>
				<releases>
					<enabled>true</enabled>
				</releases>
				<snapshots>
					<enabled>false</enabled>
				</snapshots>
			</repository>
		</repositories>

and then add the following dependency:

		<dependency>
			<groupId>uk.co.certait</groupId>
			<artifactId>html-exporter</artifactId>
			<version>0.3</version>
		</dependency>
    
Further documentation to follow.

 


[![Bitdeli Badge](https://d2weczhvl823v0.cloudfront.net/alanhay/html-exporter/trend.png)](https://bitdeli.com/free "Bitdeli Badge")

