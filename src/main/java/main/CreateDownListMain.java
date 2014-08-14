package main;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Workbook;

import createDownList.dto.CreateDownListDto;
import createDownList.service.CreateDownListService;
public class CreateDownListMain {
	public static void main(String[] args) {
		String inputFileName= JOptionPane.showInputDialog("請輸入檔案名稱，包含副檔名(直接按Enter離開)");
		//沒做副檔名的驗證
		if(inputFileName == null || inputFileName.equals("")){
			JOptionPane.showMessageDialog(null, "離開");
			System.exit(0);
		}
		String inputDataCount= JOptionPane.showInputDialog("請輸入資料數量(直接按Enter離開)");
		if(inputDataCount == null || inputDataCount.equals("")){
			JOptionPane.showMessageDialog(null, "離開");
			System.exit(0);
		}
		int dataCount=Integer.parseInt(inputDataCount);
		
		String filename = System.getProperty("user.dir")
				+ "/CreateDownListExcel/"+inputFileName;
//		String filename = System.getProperty("user.dir")
//				+ "/CreateDownListExcel/testValidateExcelDownList20140527.xls";
		CreateDownListDto param = new CreateDownListDto();
		CreateDownListService service = new CreateDownListService();
		param.setDataCount(dataCount);
		param.setFilename(filename);
		Workbook workbook = service.getWorkbook(param);
		writeFile(workbook, filename);
		JOptionPane.showMessageDialog(null, "Excel產生完畢");
	}

	private static void writeFile(Workbook workbook, String filename) {
		// 寫檔案
		FileOutputStream fso;
		try {
			fso = new FileOutputStream(filename);
			workbook.write(fso);
			fso.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
