# Excel palette for BeepBeep 3

This repository contain a palette (extension) of [BeepBeep3](https://liflab.github.io/beepbeep-3/) for the manipulation of Excel data as events streams. All projects require `beepbeep-3.jar` in their classpath (or alternately, must point to the Core source files from BeepBeep's repository) in order to compile and run.

The palette now contain the following files :

- `ExcelReader.java`: It's the main file of the palette, containing constructors and methods for the extraction of data.

- `ExcelReaderExample.java`: Create an Excel file with numbers in, and, in the main method, call an ExcelReader processor to extract and show the data of this file.

- `ExcelReaderExampleTest.java`: Some JUnit tests of ExcelReaderExample.java.

- `ExcelReaderTest.java`: Some JUnit tests of ExcelReader.java.

- `ExcelReaderExceptions.java`: Exceptions.

About the authors                                                  {#about}
-----------------

This palette was written by @ntaff and Valentin Ramos.
