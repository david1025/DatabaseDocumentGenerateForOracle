package tj.david.utils;

import java.io.InputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public interface HSSF2BeanConverter<T> {
	
	List<T> xls2bean(InputStream is, String sheet)throws Exception;

	List<T> xls2bean(byte[] buffer, String sheet)throws Exception;

	List<T> xls2bean(HSSFWorkbook workbook, String sheetname) throws Exception;
	
	
	public Boolean getFormula() ;

	public void setFormula(Boolean formula);

	public void setMatchFieldNames(String[] match);

	public void setStatEndNum(Integer start, Integer end);
	
	public HSSFWorkbook bean2xls(HSSFWorkbook workbook, String sheetname,List<T> list) throws Exception ;
	
	
	

}
