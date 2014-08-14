package createDownList.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import createDownList.dto.CreateDownListDto;

import util.PoiUtil;


public class CreateDownListService {

	private static final String SHEET_NAME_HIDE = "hideSheet";
	private static final String SHEET_NAME_DATA = "DataSheet";
	private static final String[] CURRENCY_NAMES = { "台幣", "美金", "陸幣" };
	private static final Double[] CURRENCY_RATIOS = { 1.0, 30.0, 5.0 };
	private static final String[] TITLE_NAMES = { "資料項目", "輸入金額", "貨幣", "匯率", "計算結果"  };
	
	private static void createRow(Row currentRow, Object[] valueList) {

		if (valueList != null && valueList.length > 0) {
			int i = 0;
			for (Object cellValue : valueList) {
				Cell cell = currentRow.createCell(i++);
				PoiUtil.setCellValue(cell, cellValue);
			}
		}
	}

	private void setHideSheet(Workbook book,Sheet hideSheet) {
//		Sheet hideSheet = book.createSheet(SHEET_NAME_HIDE);
		// 設定下拉選單中的數值
		Row nameRow = hideSheet.createRow(0);
		createRow(nameRow, CURRENCY_NAMES);
		// 設定下拉選單中 貨幣的匯率
		Row valueRow = hideSheet.createRow(1);
		createRow(valueRow, CURRENCY_RATIOS);

	}

	private void setDataSheet(Workbook book,Sheet dataSheet,int dataCount) {
//		Sheet dataSheet = book.createSheet(SHEET_NAME_DATA);
		// 設定驗證條件
		String formilaString = SHEET_NAME_HIDE + "!$" + PoiUtil.toCol(0)
				+ "$1:$" + PoiUtil.toCol(CURRENCY_NAMES.length - 1) + "$1";
		int validateColumn = 2;
		DataValidation dataValidation = PoiUtil.getDataValidation(dataSheet,
				formilaString, 1, 1 + dataCount, validateColumn,
				validateColumn);
		dataValidation.createPromptBox("下拉選單提示", "請使用下拉選單選擇貨幣！");
		dataValidation.createErrorBox("選擇錯誤提示", "輸入的值不在列表中！");
		dataSheet.addValidationData(dataValidation);
		
		// 開始設定使用者使用的sheet的內容
		// 設定標題
		Row row = dataSheet.createRow(0);
		createRow(row, TITLE_NAMES);

		// 設定內容
		for (int i = 1; i < 1 + dataCount; i++) {
			row = dataSheet.createRow(i);
			// 資料項目
			Cell tempCell = row.createCell(0);
			tempCell.setCellValue("項目" + i);
			// 輸入金額
			tempCell = row.createCell(1);
			tempCell.setCellValue(0);
			// 貨幣(這裡直接預設第一個數值)
			tempCell = row.createCell(2);
			tempCell.setCellValue(CURRENCY_NAMES[0]);
			// 設定取得匯率的公式
			tempCell = row.createCell(3);
			tempCell.setCellFormula("LOOKUP(C" + (i + 1) + ","
					+ SHEET_NAME_HIDE + "!$"+PoiUtil.toCol(0)+"$1:$"+PoiUtil.toCol(2)+"$1,hideSheet!$"+PoiUtil.toCol(0)+"$2:$"+PoiUtil.toCol(2)+"$2)");
			// 設定取得計算結果的公式
			tempCell = row.createCell(4);
			tempCell.setCellFormula(PoiUtil.toCol(1) + (i + 1) + "*"+PoiUtil.toCol(3) + (i + 1));
		}
	}

	public Workbook getWorkbook(CreateDownListDto param) {
		Workbook workbook = PoiUtil.getNewWorkbook(param.getFilename());

		/**
		 * 這裡建立兩個sheet， 第一個是一般使用者使用的sheet 第二個是儲存下拉選單資料的隱藏sheet
		 * 
		 * 第一個sheet使用隱藏sheet來設定下拉選單的內容和下拉選單代表的數值
		 */

		Sheet dataSheet = workbook.createSheet(SHEET_NAME_DATA);
		Sheet hideSheet = workbook.createSheet(SHEET_NAME_HIDE);

		setHideSheet(workbook, hideSheet);
		setDataSheet(workbook, dataSheet,param.getDataCount());
		// 隱藏sheet
		workbook.setSheetHidden(workbook.getSheetIndex(SHEET_NAME_HIDE), true);
		// workbook.setSheetHidden(1, true);
		return workbook;

	}
}
