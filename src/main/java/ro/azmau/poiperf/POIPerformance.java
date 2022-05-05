package ro.azmau.poiperf;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class POIPerformance {
	private static final Log	log			= LogFactory.getLog(POIPerformance.class);

	private static final String	EXPORT_FILE	= "test-poi-performance.xlsx";
	private static final int	COLUMNS		= 30;
	private static final int	ROWS		= 2_000;

	public static void main(String[] args) {
		List<Record> data = generateData(ROWS, COLUMNS);

		// Excel export duration in Java 8: 0m:19s:219ms:006545ns
		// Excel export duration in Java 11: 2m:19s:537ms:521560ns
		exportToExcel(data);
	}

	private static void exportToExcel(List<Record> data) {
		System.setProperty("org.apache.poi.util.POILogger", "org.apache.poi.util.CommonsLogger");
		long startTime = System.nanoTime();
		SXSSFWorkbook workbook = new SXSSFWorkbook();
		SXSSFSheet sheet = workbook.createSheet();

		// computing column width is the heavy operation
		sheet.trackAllColumnsForAutoSizing();

		fillBody(data, sheet);
		log.info(String.format("Excel export duration: %s", formatDuration(System.nanoTime() - startTime)));

		try (FileOutputStream fos = new FileOutputStream(EXPORT_FILE)) {
			workbook.write(fos);
		}
		catch (IOException e) {
			log.error("", e);
		}

		try {
			workbook.close();
		}
		catch (IOException e) {
			log.error("Exception on closing workbook", e);
		}
	}

	private static void fillBody(List<Record> data, SXSSFSheet sheet) {
		for (Record record: data) {
			SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
			fillRow(row, record);
		}
	}

	private static void fillRow(SXSSFRow row, Record record) {
		for (int i = 1; i < COLUMNS + 1; i++) {
			SXSSFCell cell = row.createCell(i - 1, CellType.STRING);
			String value = null;
			try {
				value = (String)PropertyUtils.getSimpleProperty(record, "col" + i);
			}
			catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
				log.error("", e);
			}
			cell.setCellValue(value);
		}
	}

	private static List<Record> generateData(int rows, int cols) {
		List<Record> result = new ArrayList<>();
		for (int i = 1; i < rows + 1; i++) {
			Record row = new Record();
			for (int j = 1; j < cols + 1; j++) {
				try {
					PropertyUtils.setSimpleProperty(row, "col" + j, String.format("Row %4d Column %d", i, j));
				}
				catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
					log.error("", e);
				}
			}
			result.add(row);
		}
		return result;
	}

	private static String formatDuration(long durationNanos) {
		long minutes = durationNanos / 60_000_000_000L;

		long undermin = durationNanos % 60_000_000_000L;
		long seconds = undermin / 1_000_000_000L;

		long undersec = undermin % 1_000_000_000L;
		long milis = undersec / 1_000_000L;
		long nanos = undersec % 1_000_000L;
		return String.format("%dm:%2ds:%03dms:%06dns", minutes, seconds, milis, nanos);
	}

}
