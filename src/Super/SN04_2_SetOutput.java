package Super;

import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IIncorporateFileEventInfo;

import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;
import jxl.Sheet;
import jxl.StringFormulaCell;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import record.Log;
import common.SN04_2_CommonUtil;
import common.SN04_2_GetInf;
import util.AUtil;
import util.Ini;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.security.KeyStore.LoadStoreParameter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import javax.swing.RowFilter;

public abstract class SN04_2_SetOutput {

	static String wTarget; // 標示要使用的 Excel
	protected static int errorCount;
	protected IAgileSession admin;
	protected String FILEPATH;
	protected boolean failuresetname = false;// 若都沒失敗就是false
	protected String failuremessage = "";
	boolean failure = false;
	boolean error = false;
	static HashMap<String, HashMap<String, ArrayList<Cell>>> itemConfig;// 用於存放 Excel 中的規則

	public HashMap<String, HashMap<String, ArrayList<Cell>>> getexcelMap() {
		return itemConfig;
	}

	public boolean getfailuresetname() {
		// 若有跑錯(通常是配方) 此時會回傳true
		return failuresetname;
	}

	public String getfailuremessage() {
		return failuremessage;
	}

	public void setfailuremessage(String errorMessage) {
		failuremessage += errorMessage + "\n\r";
		errorCount += 1;
	}

	public void resetfailuremessage() {
		failuremessage = "";
	}

	public void resetgetfailuresetname() {
		failuresetname = false;
	}

	public int getErrorCount() {
		return errorCount;
	}

	public void resetCount() {
		errorCount = 0;
	}

	public void action(IChange change, IAgileSession admin, Ini ini, Log logger) throws Exception {
		try {
			change = (IChange) admin.getObject(IChange.OBJECT_TYPE, change.getName());
			IAgileClass iAgileClass = change.getAgileClass();
			// IChange changeOrder = (IChange) change;
			logger.log(1, "Get Change as Admin:" + change);
			ITable affectedTable = SN04_2_GetInf.getAffectedTable(change);
			ITable.ISortBy sortBy = affectedTable.createSortBy(
					iAgileClass.getAttribute(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_TYPE),
					ITable.ISortBy.Order.DESCENDING);
			ITwoWayIterator iterator = affectedTable.getTableIterator(new ITable.ISortBy[] { sortBy });
			IRow row;
			IItem item;

			// Create Item Rule Map
			ArrayList<IItem> itemlist = new ArrayList<>();
			iterator = affectedTable.getTableIterator(new ITable.ISortBy[] { sortBy });
			while (iterator.hasNext()) {
				row = (IRow) iterator.next();
				// 跳過廠區有值
				if (!row.getValue(12208).toString().equals(""))
					continue;
				item = (IItem) row.getReferent();
				itemlist.add(item);
			}

			// Get Target Workbook
			wTarget = getExcelTarget();
			// Complete itemConfig
			getexcel(wTarget, itemlist, ini, logger);
			// check map
			Iterator iter = itemConfig.keySet().iterator();
			while (iter.hasNext()) {
				String key = (String) iter.next();
				// TODO arraylist datatype change
				HashMap<String, ArrayList<Cell>> val = itemConfig.get(key);
				if (val == null)
					setfailuremessage("Excel 維護錯誤 - " + key + " 找不到對應列!!!");
			}
			if (!getfailuremessage().equals(""))
				return;
			// Catch Table again
			iterator = affectedTable.getTableIterator(new ITable.ISortBy[] { sortBy });
			// loop through affected Items
			while (iterator.hasNext()) {
				error = false;
				// get affected item
				row = (IRow) iterator.next();
				// 跳過廠區有值
				if (!row.getValue(12208).toString().equals(""))
					continue;
				item = (IItem) row.getReferent();
				logger.log(1, "物件: " + item);
				// 代入在同一張表單底下半成品料件的成品料號欄位相對的成品Redline料件編號
				IItem iItem = (IItem) row.getReferent();
				if (iItem.getCell(1539).toString().substring(0, 1).equals("2")) {
					Boolean check = searchType(affectedTable, iItem, logger);
				}
				// parse definition for excel class
				String autoOutput = getAutoOutput(item, admin, ini, logger);
				if (autoOutput.contains("順序1")) {
					error = true;
					continue;
				}
				// Check Error Value
				if (autoOutput.equals("")) {
					error = true;
					errorCount++;
					continue;
				}
				if (autoOutput.contains("###$$$")) {
					error = true;
					continue;
				}
				// check error
				if (!error) {
					failure = setOutputValue(row, row.getReferent(), autoOutput, ini, logger);
					if (failure)
						continue;
				}
			}
		} catch (

		APIException e) {
			e.printStackTrace();
			throw e;
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
	}

	private Boolean searchType(ITable affectedTable, IItem iItem, Log logger) throws APIException {
		Boolean check = false;
		if (iItem.getValue(1562).toString().toUpperCase().startsWith("TEMP")) {
			Iterator it = affectedTable.getTableIterator();
			while (it.hasNext()) {
				IRow row = (IRow) it.next();
				if (row.getReferent().toString().equals(iItem.getValue(1562).toString())) {
					IItem item = (IItem) row.getReferent();
					ITable redLineTitleBlock = item.getTable(ItemConstants.TABLE_REDLINETITLEBLOCK);
					IRow redlineRow = (IRow) redLineTitleBlock.iterator().next();
					String cell = redlineRow.getValue(ItemConstants.ATT_TITLE_BLOCK_NUMBER).toString();
					iItem.setValue(1599, cell);
					check = true;
				}
			}
		} else {
			iItem.setValue(1599, iItem.getValue(1562).toString());
			check = true;
		}
		return check;

	}

	protected abstract String getExcelTarget();

	protected String getAutoOutput(IItem item, IAgileSession session, Ini ini, Log logger)
			throws APIException, Exception {
		// get agile class
		String agileAPIName = item.getAgileClass().getAPIName();
		logger.log(1, "搜索" + agileAPIName + "對應的規則");
		HashMap map = itemConfig.get(agileAPIName);
		// 取得對應分類的excel編碼資訊
		ArrayList<Cell> cellList = SN04_2_CommonUtil.checkMinorCategory((IDataObject) item, map, logger);
		// 根據excel編碼資訊產出編碼
		if (cellList == null) {
			setfailuremessage("Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行,順序1");
			return "Excel維護錯誤 - " + item.getAgileClass().getAPIName() + ",順序1";
		}
		String autoOutput = parseRule(cellList, item, session, ini, logger);
		String target = getExcelTarget();
		logger.log(1, target + ": " + autoOutput);
		return autoOutput;
	}

	protected abstract String parseRule(ArrayList<Cell> cellList, IItem item, IAgileSession session, Ini ini,
			Log logger) throws APIException, Exception;

	protected abstract boolean setOutputValue(IRow row, IDataObject ido, String autoOutput, Ini ini, Log logger)
			throws APIException;

	private HashMap getexcel(String fileTag, ArrayList<IItem> itemlist, Ini ini, Log logger) throws APIException {
		InputStream ExcelFileToRead = null;
		Workbook wbook = null;
		String EXCEL_FILE = "";
		logger.log(">Prepare To Read Excel");
		switch (fileTag) {
		case "AN":
			EXCEL_FILE = ini.getValue("File Location", "EXCEL_FILE_PATH_NUMBER");
			break;
		case "DES":
			EXCEL_FILE = ini.getValue("File Location", "EXCEL_FILE_PATH_DESCRIPTION_STANDARD");
			break;
		case "STD":
			EXCEL_FILE = ini.getValue("File Location", "EXCEL_FILE_PATH_DESCRIPTION_STANDARD");
			break;
		}

		try {
			ExcelFileToRead = new FileInputStream(EXCEL_FILE);
			wbook = Workbook.getWorkbook(ExcelFileToRead);
			ExcelFileToRead.close();
		} catch (FileNotFoundException e) {
			logger.log("找不到該檔案，請檢查Config.ini!");
			setfailuremessage("找不到該檔案，請檢查Config.ini!");
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		Sheet sheet = null;
		switch (fileTag) {
		case "AN":
			sheet = wbook.getSheet(0);
			break;
		case "DES":
			sheet = wbook.getSheet("品名");
			break;
		case "STD":
			sheet = wbook.getSheet("規格");
			break;
		}
		return createMap(sheet, itemlist, ini, logger);
	}

	private HashMap createMap(Sheet sheet, ArrayList<IItem> itemlist, Ini ini, Log logger) throws APIException {

		itemConfig = new HashMap<>();
		HashMap<String, ArrayList<Cell>> partData = null;
		for (int count = 0; count < itemlist.size(); count++) {

			IItem item = itemlist.get(count);
			String agileAPIName = item.getAgileClass().getAPIName();

			for (int i = 1; i < sheet.getRows(); i++) {
				Cell cell1 = sheet.getCell(1, i);// Excel API

				if (cell1.getType() == CellType.EMPTY)
					continue;
				LabelCell labelCell = (LabelCell) cell1;
				String classAPIName = labelCell.getString();
				if (classAPIName.toLowerCase().toString().equals("end"))
					break;
				Cell cell2 = sheet.getCell(2, i);// 分類依據欄位
				if (classAPIName.toLowerCase().contains((agileAPIName.toLowerCase()))) {

					if (itemConfig.get(classAPIName) != null)
						partData = itemConfig.get(classAPIName);
					else
						partData = new HashMap<String, ArrayList<Cell>>();

					if (wTarget.equals("AN")) {
						// 自動編碼無判斷分類欄位
						ArrayList<Cell> cellList = getExcelCell(wTarget, sheet, i);
						partData.put("", cellList);
						itemConfig.put(agileAPIName, partData);
					} else {
						// 未來會根據cell2填寫內容判斷分類
						if (cell2.getType() != CellType.EMPTY) {
							String cell2Data = cell2.getContents().trim();
							ArrayList<Cell> cellList = getExcelCell(wTarget, sheet, i);
							// 空字串即無須判斷直接使用
							if ("".equals(cell2Data))
								partData.put("", cellList);
							else {
								partData.put(cell2Data, cellList);
							}
						} else {
							// 空字串即無須判斷直接使用
							ArrayList<Cell> cellList = getExcelCell(wTarget, sheet, i);
							partData.put("", cellList);
						}
						itemConfig.put(agileAPIName, partData);
					}
				}
			}
		}
		return itemConfig;
	}

	private ArrayList<Cell> getExcelCell(String target, Sheet sheet, int row) {

		if (row >= sheet.getRows())
			return null;
		Cell[] cellList = sheet.getRow(row);
		if (target.equals("AN")) {
			ArrayList<Cell> list = new ArrayList<>();
			for (int i = 2; i < cellList.length; i++) {
				list.add(cellList[i]);
			}
			return list;
		} else {
			ArrayList<Cell> list = new ArrayList<>();
			for (int i = 3; i < cellList.length; i++) {
				list.add(cellList[i]);
			}
			return list;
		}
	}
}
// subclass name probably can't be found due to chinese character vs apiname