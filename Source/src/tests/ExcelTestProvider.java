package tests;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Encapsule un fournisseur de test sur les données de type Excel
 */
public interface ExcelTestProvider {
    /**
     * Permet de re-serialisé un Workbook
     *
     * @param wb : le workbook à re-serialisé
     * @return Retourne le workbook re-serialisé
     */
    Workbook writeOutAndReadBack(Workbook wb);

    /**
     * Permet de charger un Workbook depuis un fichier
     *
     * @param fileName : nom du fichier à charger
     * @return Retourne une instance du Workbook chargé depuis le fichier
     */
    Workbook openSampleWorkbook(String fileName);

    /**
     * Permet de créer un workbook
     * @return Retourne une instance du workbook
     */
    Workbook createWorkbook();

    /**
     * @return Retourne "xls" ou "xlsx"
     */
    String getStandardFileNameExtension();
}
