------------------------------------------------------------------------------

OVERVIEW

iQSmartSheet allows template based Excel spreadsheet generation directly
from the command line. It takes a JSON-input file of raw data and an Excel
template file as input. From both files a new output Excel file is generated.  

iQSmartSheet implements a simple and straight forward command line wrapper
for XLSTransformer class of the powerful JXLS 1.0.6 library. JXLS itself
makes use of Apache POI. Thus the installation only needs a working java
runtime environment and can be done on Unix, Windows or Mac without any
Microsoft Office software.

JXLS supports an easy to use tag-based language to place substitution
variables and more complex operations into an Excel template.

Weblink with many examples: http://jxls.sourceforge.net/1.x/

------------------------------------------------------------------------------

PRERQUISITES

In order to initally build iQSmartSheet and download all dependend libraries
from the maven repository you will need a working installation of the
following frameworks on your system:

- Java jdk (tested with jdk V1.8.0_152)
- Maven (tested with apache-maven-3.5.2)

Set JAVA_HOME and make sure that java and maven bin-folders are added to the
execution path.

To run iQSmartSheet you only need a working installation of the Java run-time
environment.

------------------------------------------------------------------------------

BUILD

1. Open a command shell

2. Goto iQSmartSheet directory (includes file "pom.xml")

3. To build the package, simply type:

mvn install

In this step a lib-Folder is created, the main-jar is built and all
dependencies are downloaded from the maven repository to the new lib-folder.

4. To tidy up the build folders (the lib folder will remain), simply type:

mvn clean

------------------------------------------------------------------------------

INSTALL

Simply copy the lib-folder to any place you like and pass it as
classpass if you call iQSmartSheet with java (see TEST).

------------------------------------------------------------------------------

TEST

Go to the example folder:

cd example

Examine the "data.json" and "templ.xlsx" sample files. Next Type:

java -cp ..\lib\* de.iqbis.iQSmartSheet.Main data.json templ.xlsx result.xlsx

where -cp defines the path how to find the jar-files that include the
software. You may set the classpass environment variable instead.

A new file "result.xlsx" is generated from the pair of sample files.

------------------------------------------------------------------------------

JSON INPUT DATA FORMAT

The format of the JSON input data is:

[
  {
    "query": "<query-name-1>",
    "columns": [ "<col-name-1>", "<col-name-2>", "<col-name-3>", ... ],
    "types": [ "<col-type-1>", "<col-type-2>", "<col-type-3>", ... ],
    "rows": [
      [ "data-row-1-col-1", "data-row-1-col-2", "data-row-1-col-3", ... ],
      [ "data-row-2-col-1", "data-row-2-col-2", "data-row-2-col-3", ... ],
      ...
    ]
  },
  {
    "query": "<query-name-2>",
    ...
  },
  ...
]

Valid column types are:

- "NUMBER" (corresponds to java type Double for numerical values)
- "LONG"   (corresponds to java type Long for integer values)
- "DATE"   (corresponds to java type Date, expects date format "yyyy-MM-dd")

These types will automatically be converted to an appropriate Excel type.
All other type names are interpreted as string, although it is recommended to
use "STRING" to make this explicitly obvious from the JSON data file.

------------------------------------------------------------------------------

LICENSE

This project is licensed under the MIT License.
See the LICENSE.txt file for details.

------------------------------------------------------------------------------

TROUBLESHOUTING

If you encounter undefined variable warnings like:

WARNING: ![0,16]: 'user_JxLsC_.name;' undefined variable user_JxLsC_.name

Please check if all column names in the JSON input file and the Excel template
are the same (case sensitive!).

E.g. column "Name" in query "user" vs. tag name "user.name"

----------------------------------------

If you get out-of-memory errors like:

java.lang.OutOfMemoryError

please raise the java -Xms and -Xmx default values when you are running the
program, e.g.:

-Xms512m: allows java to use at least 512MB of heap memory.
-Xmx2g:   allows up to 2GB of heap memory, you may raise this value if your
          system has enough memory and you want better performance for large
          files

For futher information on java command line options visit:
https://docs.oracle.com/javase/8/docs/technotes/tools/windows/java.html

----------------------------------------

If you want to make Excel to recalculate all formulas of your generated
document as soon as the document is opened by the user you can use an POI
XSSFSheet-command. Place the following text in an unused cell anywhere in
your Excel template:

${sheet.setForceFormulaRecalculation(true)}

This should work at least for Excel documents in xlsx- and xlsm-format.

Further information on this topic:
http://jxls.sourceforge.net/1.x/reference/sheetoperations.html
https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFSheet.html
https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFWorkbook.html

------------------------------------------------------------------------------

AUTHOR

Copyright (c) 2018 M. Koetter, iQbis consulting GmbH
Lahnstr. 35, 45478 Mülheim an der Ruhr, Germany
www.iqbis.de

------------------------------------------------------------------------------

ACKNOWLEDGMENTS

http://jxls.sourceforge.net/1.x/
https://poi.apache.org/

------------------------------------------------------------------------------
