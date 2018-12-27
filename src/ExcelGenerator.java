import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * 
 * @author Shubhasish
 */
public class ExcelGenerator {

	public void generateExcelDocument(List<? extends Object> coreData) {

		@SuppressWarnings("unchecked")
		List<List<String>> inputData = (List<List<String>>) coreData;
		//Object[][] inputDataArr = coreData.toArray(new Object[inputData.size()][inputData.get(0).size()]);
		
		try {
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("Result Sheet");

			int rowNum = 0;
			System.out.println("Creating excel sheet...");
			for (int j=0;j<inputData.size();j++) {
				List<String> dataRow = inputData.get(j);
				Row row = sheet.createRow(rowNum++);
				int colNum = 0;
				String[] dataArray = dataRow.toArray(new String[dataRow.size()]);
				for(int i=0;i<dataArray.length;i++) {
					Cell cell = row.createCell(colNum++);
					cell.setCellValue(dataArray[i]);
				}
			}

			FileOutputStream outputStream = new FileOutputStream("output.xls");
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();

			System.out.println("Done");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
