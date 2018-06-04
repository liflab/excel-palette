package ca.uqac.lif.cep.excelReader;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import basic.PipingUnary.Doubler;
import ca.uqac.lif.cep.Connector;
import ca.uqac.lif.cep.Pullable;

/**
 * Permet de créer un fichier Excel rempli de valeurs numériques pour tester le processeur ExcelReader
 **/

public class ExcelReaderTests
{
  @SuppressWarnings("resource")

  public static void main(String[] args) throws Exception
  {

    // Constantes permettant de modifier le nombre de lignes et de colonnes du
    // fichier
    final int nbLignes = 10;
    final int nbColonnes = 5;

    // On créer un nouveau classeur
    Workbook wb = new HSSFWorkbook();

    // On créer une nouvelle feuille
    Sheet sheet = wb.createSheet("Nouvelle Feuille");

    // Compteur de cellules
    int k = 0;

    // On crée X lignes
    for (int i = 0; i < nbLignes; i++)
    {

      // On créer une nouvelle ligne
      Row row1 = sheet.createRow(i);

      // On créer X colonnes
      for (int j = 0; j < nbColonnes; j++)
      {

        // On incrémente le compteur de cellule avant chaque création de cellule
        k++;

        // On créer une cellule et on lui ajoute du contenu
        row1.createCell(j).setCellValue(k);

      } // On ferme le second for
    } // On ferme le premier for

    ExcelReaderExceptions ExcelExceptions = null;
    // On écrit le contenu dans le fichier de sortie
    try {OutputStream fileOut = new FileOutputStream(
        "C:\\Users\\Taffoureau\\Music\\Excel Tests\\workbook.xls");
   
          {
            wb.write(fileOut);
          } 
    
        } catch (ExcelReaderExceptions e) {
          ExcelExceptions = e;
          throw e;
        }
    finally {
    // Objets permettant de formatter le contenu d'une cellule
    DataFormatter formatter = new DataFormatter();

    // On récupère la feuille courante
    Sheet sheet1 = wb.getSheetAt(0);

    // On parcoure les colonnes
    for (Row row1 : sheet1)
    {

      // On parcoure les lignes
      for (Cell cell : row1)
      {

        // On récupère la localisation de la cellule
        CellReference cellRef = new CellReference(row1.getRowNum(), cell.getColumnIndex());

        // On affiche la localisation de la cellule
        System.out.print(cellRef.formatAsString());

        // Pour séparer la localisation de la cellule de son contenu
        System.out.print(" - ");

        // On recupère le contenu brut de la cellule
        String text = formatter.formatCellValue(cell);

        // On affiche ce contenu
        System.out.println(text);

      } // On ferme le second for
    } // On ferme le premier for
    

    ExcelReader test = new ExcelReader("C:\\Users\\Taffoureau\\Music\\Excel Tests\\workbook.xls",0,2,3);


    Doubler doubler = new Doubler();
    Connector.connect(test, doubler);
    Pullable p = doubler.getPullableOutput();

    for (int i = 0; i < 20; i++)
    {
      int x = (Integer) p.pull();

      // On affiche à l'écran
      System.out.println("Le fichier contient: " + x);

    } // On ferme le for
    
    } // On ferme le finally
  } // On ferme le main
} // On ferme la classe
