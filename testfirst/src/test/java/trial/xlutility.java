package trial;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlutility {

	String path;

	public xlutility(String path) {

		this.path = path;

	}

	FileInputStream file;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;

	public int gettotlCountRow(String sheetname) throws IOException {

		file = new FileInputStream(path);

		workbook = new XSSFWorkbook(file);

		sheet = workbook.getSheet(sheetname);

		int Count = sheet.getLastRowNum();

		file.close();
		workbook.close();

		return Count;

	}

	public int gettotalCountColum(String sheetname) throws IOException {
		file = new FileInputStream(path);
		workbook = new XSSFWorkbook(file);
		sheet = workbook.getSheet(sheetname);

		row = sheet.getRow(1);

		int countcolumn = row.getLastCellNum();

		file.close();
		workbook.close();
		return countcolumn;

	}

	public String getallthevalue(String sheetname, int rownum, int columnnum) throws IOException {

		String data = "";

		file = new FileInputStream(path);
		workbook = new XSSFWorkbook(file);

		sheet = workbook.getSheet(sheetname);

		row = sheet.getRow(rownum);
		cell = row.getCell(columnnum);

		DataFormatter formate = new DataFormatter();
		try {
			data = formate.formatCellValue(cell);
		} catch (Exception e) {
			data = "";
		}

		file.close();
		workbook.close();

		return data;

	}

}
