package tests;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Common superclass for testing implementations of
 *  {@link org.apache.poi.ss.usermodel.Cell}
 */
public abstract class BaseTestExcelReader {

  protected static ExcelTestProvider _excelTestProvider;

    /**
     * @param excelTestProvider : Objet fournissant des données de test
     * @return void
     */
    public BaseTestExcelReader(ExcelTestProvider excelTestProvider) {
      _excelTestProvider = excelTestProvider;
    }
    
    @BeforeClass
    public static void setUpBeforeClass() throws Exception {
      
      final int nbRow = 5;
      final int nbColumns = 5;

      Workbook wb = _excelTestProvider.createWorkbook();
      Sheet sheet = wb.createSheet("Feuille de test");
      int k = 0;

      for (int i = 0; i < nbRow; i++)
      {
        Row row1 = sheet.createRow(i);
        for (int j = 0; j < nbColumns; j++)
        {
          k++;
          row1.createCell(j).setCellValue(k);

        } 
      } 
    }
    
    @AfterClass
    public static void setUpAfterClass() throws Exception {
      
      
    }
    
    @Before
    public static void setUpBefore() throws Exception {
      
      
    }

    @After
    public static void setUpAfter() throws Exception {
      
      
    }

   
}
