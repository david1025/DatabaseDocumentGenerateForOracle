package tj.david.utils;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

/**
 * Excel2Bean
 * @author David
 */
public abstract class HSSF2BeanAbstract<T> implements HSSF2BeanConverter<T> {

	protected Class<T> clazz;

	protected String[] MatchFieldNames; // 实体和excle对应的列

	protected Integer start; // 开始行

	protected Integer end; // 结束行

	protected Boolean formula = false; //是否是公式

	protected HSSFFormulaEvaluator hfe; // 03格式

	protected XSSFFormulaEvaluator xhfe; // 07格式

	protected Integer moveLines; // 要移动的行数
	protected Integer[] lastLines; // 哪几行需要移动

	public void setMoveLines(Integer moveLines) {
		this.moveLines = moveLines;
	}

	public void setLastLines(Integer[] lastLines) {
		this.lastLines = lastLines;
	}

	public List<T> xls2bean(InputStream is, String sheet) throws Exception {
		HSSFWorkbook workbook = new HSSFWorkbook(is);

		List<T> beanList = this.xls2bean(workbook, sheet);
		return beanList;
	}

	public List<T> xls2bean(byte[] buffer, String sheet) throws Exception {

		return this.xls2bean(new ByteArrayInputStream(buffer), sheet);
	}

	public static Object getValue(HSSFCell cell, HSSFFormulaEvaluator hfe, Boolean formula) {
		Object value = null;
		Integer ct = null;
		if (cell != null) {
			ct = cell.getCellType();

			if (ct == 0) {
				value = new Double(cell.getNumericCellValue());
			} else if (ct == 1) {
				value = cell.getStringCellValue();
			} else if (ct == 2) {
				if (formula) {
					value = cell.getCellFormula();
				} else {

					try {
						CellValue cv = hfe.evaluate(cell);
						value = cv.formatAsString();
					} catch (Exception e) {
						value = "";
					}
				}
			} else if (ct == 3) {
				value = "";
			} else if (ct == 4) {
				value = new Boolean(cell.getBooleanCellValue());
			} else if (ct == 5) {
				value = "ERROR";
			}
		}
		return value;
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

	public static XSSFCell setValue(XSSFCell cell, Object cellvalue, Boolean formula) {
		
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
