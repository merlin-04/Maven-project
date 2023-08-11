package excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	XSSFSheet sheet;
	Row row;
	Cell cell;

	ExcelRead() throws IOException {
		FileInputStream file = new FileInputStream("C:\\Users\\91623\\Desktop\\Java\\ReadExcel.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		sheet = workbook.getSheet("Sheet1");
	}

	public String readData(int i, int j) {
		row = sheet.getRow(i);
		cell = row.getCell(j);

		CellType type = cell.getCellType();
		try {
			switch (type) {
			case NUMERIC:
				double data = cell.getNumericCellValue();
				return String.valueOf(data);
			case STRING:
				return cell.getStringCellValue();
			default:

				System.out.println("INVALID");
			}

		} catch (Exception e) {

		}
		return "";

	}

	public int rowSize() {
		int rows = sheet.getLastRowNum() + 1;
		return rows;
	}

	public int cellSize(int i) {
		int cells = sheet.getRow(i).getLastCellNum();
		return cells;
	}

	public static void main(String args[]) throws IOException {
		ExcelRead e = new ExcelRead();

		for (int i = 0; i < e.rowSize(); i++) {

			for (int j = 0; j < e.cellSize(1); j++) {

				System.out.println(e.readData(i, j) + "  ");

			}
			System.out.println();
		}
	}

}
