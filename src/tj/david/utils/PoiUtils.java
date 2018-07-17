package tj.david.utils;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class PoiUtils<T> {

	protected Class<T> clazz;

	protected HSSFFormulaEvaluator hfe = null;
	protected String[] MatchFieldNames;

	protected Integer start;
	protected Integer end;

	protected Boolean formula = false;

	@SuppressWarnings("rawtypes")
	public HSSFWorkbook bean2xls(HSSFWorkbook workbook, String sheetname, List<T> list) throws Exception {

		HSSFSheet sh = workbook.getSheet(sheetname);

		if (hfe == null) {
			hfe = new HSSFFormulaEvaluator(workbook);
		}
		for (int i = 0; i < MatchFieldNames.length; i++) {
			if (MatchFieldNames[i] == null || MatchFieldNames[i].equals("")) {
				continue;
			}

			HSSFRow row = sh.getRow(i);

			for (int j = 0; j < end - start; j++) {

				T intance = list.get(j);

				Field f = clazz.getDeclaredField(MatchFieldNames[i]);

				Class type = f.getType();

				String getMethodname = "get" + MatchFieldNames[i].substring(0, 1).toUpperCase()
						+ MatchFieldNames[i].substring(1);

				Method method = clazz.getMethod(getMethodname);

				HSSFCell cell = row.getCell(j + start);
				Object value = null;

				if (Integer.class.equals(type)) {
					value = method.invoke(intance);
				}
				if (String.class.equals(type)) {
					value = method.invoke(intance);
				}
				if (Double.class.equals(type)) {
					value = method.invoke(intance);
				}
				if (value != null) {
					if (cell == null) {
						cell = row.createCell(j + start);
					}
					setValue(cell, value, formula);
				}
			}
		}
		return workbook;
	}

	public static HSSFCell setValue(HSSFCell cell, Object cellvalue, Boolean formula) {

		Integer ct = null;

		ct = cell.getCellType();
		if (ct == 0) {
			if (cellvalue.getClass().getName().equals("java.lang.Double")) {
				cell.setCellValue((Double) cellvalue);
			} else if (cellvalue.getClass().getName().equals("java.lang.Float")) {
				cell.setCellValue((Float) cellvalue);
			} else if (cellvalue.getClass().getName().equals("java.lang.Integer")) {
				cell.setCellValue(cellvalue.toString());
			} else if (cellvalue.getClass().getName().equals("java.lang.Long")) {
				cell.setCellValue(cellvalue.toString());
			}
		} else if (ct == 1) {

			if (cellvalue.getClass().getName().equals("java.util.Date")) {
				cell.setCellValue((Date) cellvalue);
			} else {
				cell.setCellValue(cellvalue.toString());
			}
		} else if (ct == 2) {
			if (formula) {
				cell.setCellFormula(cellvalue.toString());
			}
		} else if (ct == 3) {
			if (cellvalue.getClass().getName().equals("java.lang.Double")) {
				cell.setCellValue((Double) cellvalue);
			} else if (cellvalue.getClass().getName().equals("java.lang.Float")) {
				cell.setCellValue((Float) cellvalue);
			} else if (cellvalue.getClass().getName().equals("java.lang.Integer")) {
				cell.setCellValue(cellvalue.toString());
			} else if (cellvalue.getClass().getName().equals("java.lang.Long")) {
				cell.setCellValue(cellvalue.toString());
			} else if (cellvalue.getClass().getName().equals("java.lang.String")) {
				cell.setCellValue(cellvalue.toString());
			}
		} else if (ct == 4) {
			cell.setCellValue(Boolean.parseBoolean(cellvalue.toString()));
		} else if (ct == 5) {
		}

		return cell;

	}
}
