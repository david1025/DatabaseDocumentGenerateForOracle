package tj.david.utils;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HSSF2BeanColumMatch<T> extends HSSF2BeanAbstract<T> {

	public HSSF2BeanColumMatch(Class<T> clazz) {
		this.clazz = clazz;
	}

	public HSSF2BeanColumMatch(Class<T> clazz, Integer startrow, Integer endrow, String[] ColumNumMatchFieldNames) {

		this.clazz = clazz;
		this.MatchFieldNames = ColumNumMatchFieldNames;
		this.end = endrow;
		this.start = startrow;

	}

	public Boolean getFormula() {

		return formula;
	}

	public void setFormula(Boolean formula) {
		this.formula = formula;
	}

	public void setMatchFieldNames(String[] match) {

		this.MatchFieldNames = match;
	}

	public void setStatEndNum(Integer start, Integer end) {
		this.start = start;
		this.end = end;
	}

	/**
	 * 将excle转化成Bean
	 * @MethodName:xls2bean
	 * @author: 张大伟
	 * @email: zhangdawei@91isoft.com 
	 * @date 2016年5月30日 下午10:34:37
	 * @version V1.0
	 * @param workbook
	 * @param sheetname
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("rawtypes")
	public List<T> xls2bean(HSSFWorkbook workbook, String sheetname) throws Exception {
		List<T> beanList = new ArrayList<T>();

		HSSFSheet sheet = workbook.getSheet(sheetname);
		if (hfe == null) {
			hfe = new HSSFFormulaEvaluator(workbook);
		}
		for (int i = start; i < end; i++) {
			HSSFRow row = sheet.getRow(i);

			if (row == null) {
				continue;
			}

			T obj = null;

			obj = clazz.newInstance();
			for (int j = 0; j < MatchFieldNames.length; j++) {

				if (MatchFieldNames[j] == null || MatchFieldNames[j].equals("")) {
					continue;
				}

				HSSFCell cell = row.getCell(j);

				Field f;

				f = clazz.getDeclaredField(MatchFieldNames[j]);

				Class type = f.getType();

				String getMethodname = "set" + MatchFieldNames[j].substring(0, 1).toUpperCase()
						+ MatchFieldNames[j].substring(1);

				Method method = clazz.getMethod(getMethodname, type);

				Object value = getValue(cell, hfe, formula);

				if (value == null) {
					value = "";
				}

				if (Integer.class.equals(type)) {
					method.invoke(obj, Integer.parseInt(value.toString()));
				}

				if (String.class.equals(type)) {
					method.invoke(obj, value.toString());
				}

				if (Double.class.equals(type)) {
					if (NumberUtils.isNumber(value.toString())) {
						method.invoke(obj, Double.parseDouble(value.toString()));
					}
				}

			}

			beanList.add(obj);

		}
		return beanList;

	}

	/**
	 * 实体转换成xls（03格式）
	 * @MethodName:bean2xls
	 * @author: 张大伟
	 * @email: zhangdawei@91isoft.com 
	 * @date 2016年5月30日 下午10:33:37
	 * @version V1.0
	 * @param workbook
	 * @param sheetname
	 * @param list
	 * @return
	 * @throws Exception
	 */
	public HSSFWorkbook bean2xls(HSSFWorkbook workbook, String sheetname, List<T> list) throws Exception {
		HSSFSheet sheet = workbook.getSheet(sheetname);
		if (hfe == null) {
			hfe = new HSSFFormulaEvaluator(workbook);
		}

		if (lastLines != null && lastLines[0] < start + list.size()) {
			workbook.getSheet(sheetname).shiftRows(lastLines[0], lastLines[lastLines.length - 1], moveLines);
		}

		HSSFRow row_first = sheet.getRow(start);

		for (int i = 0; i < end - start; i++) {
			HSSFRow row = sheet.getRow(i + start);

			if (row == null) {
				row = sheet.createRow(i + start);
			}

			T obj = list.get(i);

			for (int j = 0; j < MatchFieldNames.length; j++) {

				if (MatchFieldNames[j] == null || MatchFieldNames[j].equals("")) {
					continue;
				}
				HSSFCell cell = row.getCell(j);

				if (cell == null) {
					cell = row.createCell(j);
				}
				cell.setCellStyle(row_first.getCell(j).getCellStyle());

				String getMethodname = "get" + MatchFieldNames[j].substring(0, 1).toUpperCase()
						+ MatchFieldNames[j].substring(1);
				Method method = clazz.getMethod(getMethodname);
				Object value = method.invoke(obj);
				if (value != null) {
					setValue(cell, value, formula);
				}
			}
		}

		return workbook;
	}

	/**
	 * 实体转换成xlsx（07格式）
	 * @MethodName:bean2xlsx
	 * @author: 张大伟
	 * @email: zhangdawei@91isoft.com 
	 * @date 2016年5月30日 下午10:32:16
	 * @version V1.0
	 * @param workbook excle
	 * @param sheetname sheetname名
	 * @param list 要写入的数据
	 * @return
	 * @throws Exception
	 */
	public XSSFWorkbook bean2xlsx(XSSFWorkbook workbook, String sheetname, List<T> list) throws Exception {
		XSSFSheet sheet = workbook.getSheet(sheetname);
		if (xhfe == null) {
			xhfe = new XSSFFormulaEvaluator(workbook);
		}

		if (lastLines != null && lastLines[0] < start + list.size()) {
			workbook.getSheet(sheetname).shiftRows(lastLines[0], lastLines[lastLines.length - 1], moveLines);
		}

		XSSFRow row_first = sheet.getRow(start);
		
		if(row_first == null){
			row_first = sheet.createRow(start);
		}

		for (int i = 0; i < end - start; i++) {
			XSSFRow row = sheet.getRow(i + start);

			if (row == null) {
				row = sheet.createRow(i + start);
			}

			T obj = list.get(i);

			for (int j = 0; j < MatchFieldNames.length; j++) {

				if (MatchFieldNames[j] == null || MatchFieldNames[j].equals("")) {
					continue;
				}
				XSSFCell cell = row.getCell(j);

				if (cell == null) {
					cell = row.createCell(j);
				}
				cell.setCellStyle(row_first.getCell(j).getCellStyle());

				String getMethodname = "get" + MatchFieldNames[j].substring(0, 1).toUpperCase()
						+ MatchFieldNames[j].substring(1);
				Method method = clazz.getMethod(getMethodname);
				Object value = method.invoke(obj);
				if (value != null) {
					setValue(cell, value, formula);

				}
			}
		}
		

		return workbook;
	}
	/**
	 * 
	 * setFormula:
	 * 
	 * @author 张大伟
	 * @param sheet
	 * @param formulaArray
	 * @param line
	 * @since Ver 1.1
	 */
	public void setFormula(HSSFSheet sheet, String[] formulaArray, Integer line) {
		HSSFRow row = sheet.getRow(line);

		for (int j = 0; j < formulaArray.length; j++) {
			HSSFCell cell = row.getCell(j);
			if (formulaArray[j] != "合计" && !"".equals(formulaArray[j])) {
				cell.setCellFormula(formulaArray[j]);
				cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
			} else {

				if (cell != null) {
					cell.setCellValue(formulaArray[j]);
				}
			}
		}
	}

	public void setFormulaRow(HSSFSheet sheet, Integer cellNum, Integer star, Integer end) {
		end = end + star;
		HSSFRow r = sheet.getRow(star);
		HSSFCell c = r.getCell(cellNum);
		String formula = "";
		if (c.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
			formula = c.getCellFormula();
		}
		for (int i = star; i < end; i++) {
			String f = formula.replace((star + 1) + "", (i + 1) + "");
			HSSFRow row = sheet.getRow(i);
			if (!"".equals(f) && null != f) {
				for (int j = 0; j < row.getLastCellNum(); j++) {
					if (j == cellNum) {
						HSSFCell cell = row.getCell(j);
						cell.setCellFormula(f);
						cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
					}
				}
			}
		}
	}

	public void setFormulaRow2(HSSFSheet sheet, Integer cellNum, Integer star, Integer end) {
		end = end + star;
		HSSFRow r = sheet.getRow(star);
		HSSFCell c = r.getCell(cellNum);
		String formula = "";
		if (c.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
			formula = c.getCellFormula();
		}
		for (int i = star; i < end; i++) {
			String f = formula.replace((star + 1) + "", (end + 1) + "");
			HSSFRow row = sheet.getRow(i);
			if (!"".equals(f)) {
				for (int j = 0; j < row.getLastCellNum(); j++) {
					if (j == cellNum) {
						HSSFCell cell = row.getCell(j);
						cell.setCellFormula(f);
						cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
					}
				}
			}
		}
	}
}
