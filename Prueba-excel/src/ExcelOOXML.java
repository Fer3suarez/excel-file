import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOOXML {

    private static final Logger LOGGER = Logger.getLogger("mx.com.hash.newexcel.ExcelOOXML");
    
    private static File archivo = new File("prueba-excel.xlsx");
    private static Workbook workbook = new XSSFWorkbook();
    private static Sheet pagina = workbook.createSheet("Hoja de prueba");
    private static CellStyle style = workbook.createCellStyle();
	private static String nombreArchivo = "prueba-excel.xlsx";
	private static String rutaArchivo = "C:\\Users\\Ferna\\git\\prueba-excel\\" + nombreArchivo;
    
    public static void createExcel() {
    	
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        String[] titulos = {"Nombre", "Apellidos", "Email", "DNI"};
        String[] datos   = {"Fernando", "Suárez Menéndez", "suarez11fer@gmail.com", "00000000X"};

        Row fila = pagina.createRow(1);
        for (int i = 0; i < titulos.length; i++) {
            Cell celda = fila.createCell(i);
            celda.setCellStyle(style);
            celda.setCellValue(titulos[i]);
        }
        fila = pagina.createRow(2);
        for (int i = 0; i < datos.length; i++) {
            Cell celda = fila.createCell(i);
            celda.setCellValue(datos[i]);
        }
        //Guardar archivo
        try {
            FileOutputStream salida = new FileOutputStream(archivo);
            if (archivo.exists()) {// si el archivo existe se elimina
            	archivo.delete();
				System.out.println("Archivo eliminado");
			}
            workbook.write(salida);
            workbook.close();
            LOGGER.log(Level.INFO, "Archivo creado existosamente en {0}", archivo.getAbsolutePath());
        } catch (FileNotFoundException ex) {
            LOGGER.log(Level.SEVERE, "Archivo no localizable en sistema de archivos");
        } catch (IOException ex) {
            LOGGER.log(Level.SEVERE, "Error de entrada/salida");
        }
    }
    
    public static void readExcel() {
 
		try (FileInputStream file = new FileInputStream(new File(rutaArchivo))) {
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
 
			Row row;
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				Cell cell;
				while (cellIterator.hasNext()) {
					cell = cellIterator.next();
					System.out.print(cell.getStringCellValue()+" | ");
				}
				System.out.println();
			}
		} catch (Exception e) {
			e.getMessage();
		}
    }

    public static void main(String[] args) {
    	createExcel();
    	readExcel();
    }
}