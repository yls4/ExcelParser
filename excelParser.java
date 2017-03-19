import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;


//Import the JExcel API
import jxl.Workbook;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import jxl.Cell;

// Only works on Excel 97-2003 files.  If you receive an error, save the excel file as a 97-2003 file and try again.
public class excelParser {
	static Workbook workbook;
	static WritableWorkbook copy;
	static String wordList = "wordList.txt";
	public static void main(String[] args) throws BiffException, IOException, WriteException {
		Scanner reader = new Scanner(System.in);
		System.out.println("Enter file name");
		String fileName = reader.nextLine();
		try {
			workbook = Workbook.getWorkbook(new File(fileName));
			copy = Workbook.createWorkbook(new File("temp.xls"), workbook);
			FileReader input = new FileReader(wordList);
			BufferedReader list = new BufferedReader(input);
			String myLine = null;
			ArrayList<String> words = new ArrayList<String>();
			
			while ((myLine = list.readLine()) != null) {
				words.add(myLine);
			}
			
			//WorkSheet sheet = workbook.getSheet(1);
			WritableSheet sheet2 = copy.getSheet(0);
			int rows = sheet2.getRows();
			int cols = sheet2.getColumns();
			
			for (int row=0; row<rows; row++) {
				for (String s: words) {
					for (int col=0; col<cols; col++) {
						Cell cell = sheet2.getCell(col, row);
						if (cell.getContents().equals(s)) {
							sheet2.removeRow(row);
							continue;
						}	
					}
				}
			}
			
			copy.write();
			copy.close();
			System.out.println("Output stored in temp.xls");
		}
		catch(Error E) {
			System.out.println("File not found");
		}
	}
}
