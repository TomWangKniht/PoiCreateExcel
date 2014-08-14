package util;

import java.util.Date;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiUtil {

	public static Workbook getNewWorkbook(String fullFileName){
		String subName =FileUtil.getSubFileName(fullFileName);

		if (subName.equalsIgnoreCase(".xls")) {
			return new HSSFWorkbook();
		} else if (subName.equalsIgnoreCase(".xlsx")) {
			return new XSSFWorkbook();
		} else {
			return null;
		}
	}
	
	public static void setCellValue(Cell cell,Object cellValue){
		if (cellValue instanceof String) {
			cell.setCellValue((String) cellValue);
		} else if (cellValue instanceof Double) {
			cell.setCellValue((Double) cellValue);
		}else if (cellValue instanceof Boolean) {
			cell.setCellValue((Boolean) cellValue);
		}else if (cellValue instanceof Date) {
			cell.setCellValue((Date) cellValue);
		}
	}

	public static String toCol(int col) {	
		if (col < 26) {
			return String.valueOf((char) (65 + col));
		}
		int mainCount = col / 26;
		int minorCount = col - 26 * mainCount;
		char mainLetter = (char) (64 + mainCount);
		char minorLetter = (char) (65 + minorCount);
		return String.valueOf(mainLetter) + String.valueOf(minorLetter);

	}

	public static DataValidation getDataValidation(Sheet sheet , String formula,
			int firstRow, int lastRow, int firstCol, int lastCol){
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
//		DVConstraint constraint = DVConstraint
//				.createFormulaListConstraint(formula);
		  DataValidationConstraint constraint = dvHelper.createFormulaListConstraint(
				  formula);
		CellRangeAddressList regions = new CellRangeAddressList(
				firstRow, lastRow, firstCol,
				lastCol);
		DataValidation dataValidation = dvHelper.createValidation(constraint,
				regions);
		return dataValidation;
//		dataValidation.createPromptBox("下拉選單提示","請使用下拉選單選擇貨幣！");     
//		dataValidation.createErrorBox("選擇錯誤提示","輸入的值不在列表中！");  		
//		sheet.addValidationData(dataValidation);
		
	}
	
//	public static void setDataValidation(Sheet sheet , String formula,
//			int[] validateRange){
//		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
//		for(int row=validateRange[0];row<=validateRange[1];row++){
//			for(int col = validateRange[2];col<=validateRange[3];col++){
//				sheet.addValidationData(getDataValidation(dvHelper,formula,row,row,col,col));
//			}
//		}
//		
//	}
//	private static DataValidation getDataValidation(DataValidationHelper dvHelper ,String formula,
//			int firstRow,int lastRow,int firstCol,int lastCol) {
//
//		DVConstraint constraint = DVConstraint
//				.createFormulaListConstraint(formula);
//		CellRangeAddressList regions = new CellRangeAddressList(
//				firstRow, lastRow, firstCol,
//				lastCol);
//		
//
//		DataValidation dataValidation = dvHelper.createValidation(constraint,
//		regions);
//		
//		dataValidation.createPromptBox("下拉選單提示","請使用下拉選單選擇貨幣！");     
//		dataValidation.createErrorBox("選擇錯誤提示","輸入的值不在列表中！");  
//		return dataValidation;
//	}
}
