package AutoOutput;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import com.agile.api.APIException;
import com.agile.api.ChangeConstants;
import com.agile.api.IAgileClass;
import com.agile.api.IAgileList;
import com.agile.api.IAgileSession;
import com.agile.api.IAutoNumber;
import com.agile.api.ICell;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.IQuery;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;

//0607EDITED
import Super.SN04_2_SetOutput;
import common.SN04_2_GetInf;
import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import record.Log;
import util.AUtil;
import util.Ini;;

public class SN04_2_SetNumber extends SN04_2_SetOutput {

	public void action(IChange change, IAgileSession admin, Ini ini, Log logger) throws Exception {
		logger.log(1, "Set Number......");
		IChange changeAF = (IChange) admin.getObject(IChange.OBJECT_TYPE, change.getName());
		ITable AffectedItemsTable = changeAF.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
		Iterator it = AffectedItemsTable.iterator();
		while(it.hasNext()) {
			IRow row = (IRow) it.next();		
			IDataObject item = row.getReferent();
			ITable redLineTitleBlock = item.getTable(ItemConstants.TABLE_REDLINETITLEBLOCK);
			IRow redlineRow = (IRow) redLineTitleBlock.iterator().next();
			ICell cell = redlineRow.getCell(ItemConstants.ATT_TITLE_BLOCK_NUMBER);
			cell.setValue(cell.getOldValue());		
			IItem item1 = (IItem) item;
			item1.setValue(ItemConstants.ATT_PAGE_TWO_TEXT05, "");
		}
	
		super.action(change, admin, ini, logger);
	}

	@Override
	protected String getExcelTarget() {
		return "AN";
	}

	protected String parseRule(ArrayList<Cell> cellList, IItem item, IAgileSession admin, Ini ini, Log logger)
			throws Exception {
		String autoNumber = "";
		try {
			if (cellList == null) {
				logger.log(1, "找不到規則" + "!!!");
				return "";
			}
			// Get Row
			String error = "###$$$";
			for (int i = 0; i < cellList.size(); i++) {
				Cell cell = cellList.get(i);
				String celli = "順序" + (i + 1); // 順序01開始
				if (cell.getType() == CellType.EMPTY)
					continue;
				logger.log(1, celli + " 規則為 " + cell.getContents());
				if (cell.getType() == CellType.LABEL) {
					LabelCell labelCell = (LabelCell) cell;
					String c = labelCell.getString().trim();
					if (c.equals("end"))
						break;
					else if (c.charAt(0) == '$') {
						if (c.length() < 2) {
							autoNumber += error;
							super.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						if (!c.contains("nbsp"))
							autoNumber += c.substring(1) + "";
						else
							autoNumber += " ";
					} // end else if - '$'
					else if (c.charAt(0) == '@') {
						if (!c.contains(",")) {
							autoNumber += error;
							super.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						String[] rules = c.substring(1).split(",");
						if (rules.length != 3) {
							autoNumber += error;
							super.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						char c0 = rules[1].toLowerCase().charAt(0);
						char c1 = rules[1].toLowerCase().charAt(1);
						if (SN04_2_GetInf.check(c1)) {
							autoNumber += error;
							super.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						c0 = rules[2].toLowerCase().charAt(0);
						if (SN04_2_GetInf.check(c0)) {
							autoNumber += error;
							super.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}

						String ObjectInfo = SN04_2_GetInf.getListValue(item, logger, admin, rules[0], rules[1], rules[2]);

						if (ObjectInfo.contains("欄位為空")) {
							autoNumber += error;
							super.setfailuremessage(ObjectInfo);
						}
						if (ObjectInfo.contains("請檢查 Excel")) {
							autoNumber += error;
							super.setfailuremessage(ObjectInfo);
						}
						if (ObjectInfo.contains("維護錯誤")) {
							autoNumber += error;
							super.setfailuremessage(ObjectInfo);
						}
						autoNumber += ObjectInfo;

					} // end else if - '@'
					else if (c.charAt(0) == '~') {
						// 流水碼
						if (!c.contains(",")) {
							autoNumber += error;
							super.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						String[] rules = c.substring(1).split(",");
						if (rules[0].contains("流水碼")) {
							int rule = Integer.parseInt(rules[1]);

							// Initialize serialNumber
							String serialNumber = initializeNum(rule);
							while (checkSameNum(autoNumber + serialNumber, item, admin, ini, logger)) {
								int serialInt = Integer.parseInt(serialNumber) + 1;
								serialNumber = String.valueOf(serialInt);
								serialNumber = addZero(rule, serialNumber);
								if (checkOverNum(serialNumber, rule)) {
									autoNumber += error;
									super.setfailuremessage("流水號已滿... 請洽 Admin!!!");
								}
							}
							autoNumber += addZero(rule, serialNumber);
						}
					} // end else if - '~'

					else if (c.charAt(0) == '&') {
						String Code = c.substring(1, c.length()).replace(" ", "");
						logger.log(1, "讀取 Config ...");
						String string = ini.getValue("Excel-BookMark", Code);
						String[] config = string.split("/");
						String ObjectInfo = SN04_2_GetInf.getObjectValue(item, logger, admin, config);
						if (ObjectInfo.equals("")) {
							autoNumber += error;
							super.setfailuremessage(ObjectInfo);
						} else
							autoNumber += ObjectInfo;
					} // end else if - '&'
					else {
						String ObjectInfo = "";
						if (!c.contains(",")) {
							autoNumber += error;
							super.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						String[] rules = c.split(",");

						if (rules.length == 3) {
							char c0 = rules[1].toLowerCase().charAt(0);
							char c1 = rules[1].toLowerCase().charAt(1);
							char c2 = rules[2].toLowerCase().charAt(0);
							if (c0 != 'p') {
								autoNumber += error;
								super.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							if (SN04_2_GetInf.check(c1)) {
								autoNumber += error;
								super.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							if (SN04_2_GetInf.check(c2)) {
								autoNumber += error;
								super.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							ObjectInfo = SN04_2_GetInf.getcolunmValue(item, logger, admin, rules[0], rules[1], rules[2]);
						} else if (rules.length == 2) {
							char c0 = rules[1].toLowerCase().charAt(0);
							char c1 = rules[1].toLowerCase().charAt(1);
							if (c0 != 'p') {
								autoNumber += error;
								super.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							if (SN04_2_GetInf.check(c1)) {
								autoNumber += error;
								super.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							ObjectInfo = SN04_2_GetInf.getcolunmValue(item, logger, admin, rules[0], rules[1], null);
						} else {
							autoNumber += error;
							super.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}

						if (ObjectInfo.contains("欄位為空")) {
							autoNumber += error;
							super.setfailuremessage(ObjectInfo);
						} else if (ObjectInfo.contains("請檢查 Excel")) {
							autoNumber += error;
							super.setfailuremessage(ObjectInfo);
						} else {
							if (ObjectInfo.equals("") && autoNumber.substring(autoNumber.length() - 1).matches("\\W")) {
								autoNumber = autoNumber.substring(0, autoNumber.length() - 1);
							}else {
								autoNumber += ObjectInfo;
							}									
						}
					} // end else - no tag
				} // end if -Lable
				else if (cell.getType() == CellType.NUMBER) { // int
					NumberCell numberCell = (NumberCell) cell;
					String string = String.valueOf(numberCell.getValue()).trim();
					int c_int;
					double c_double;
					if (string.indexOf(".") > 0) {
						string = string.replaceAll("0+?$", "");// 去掉無用的0
						string = string.replaceAll("[.]$", "");// 小數點後面全為0
					}
					if (string.contains(".")) {
						c_double = Double.parseDouble(string);
						autoNumber += c_double + "";
					} else {
						c_int = Integer.parseInt(string);
						autoNumber += c_int + "";
					}
				} // end if -Number
			} // end for
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
		return autoNumber;
	}

	private static String addZero(int length, String number) {
		String tmp = "";
		for (int i = 0; i < length - number.length(); i++) {

			tmp += "0";
		}
		return tmp + number;
	}

	private boolean checkOverNum(String serialNumber, int rule) {
		int initialize = rule;
		String initializeS = "";
		for (int i = 0; i < initialize; i++)
			initializeS += "9";
		if (serialNumber.equals(initializeS))
			return true;
		else
			return false;
	}

	private String initializeNum(int rule) {
		int initialize = rule - 1;
		String initializeS = "";
		for (int i = 0; i < initialize; i++)
			initializeS += "0";
		return initializeS + "1";
	}

	private boolean checkSameNum(String number, IItem item, IAgileSession admin, Ini ini, Log logger)
			throws APIException {

		try {
			logger.log(">Checking Same Number");
			IQuery query = (IQuery) admin.createObject(IQuery.OBJECT_TYPE, item.getAgileClass().getAPIName());
			query.setCaseSensitive(false);
			// 1001 = item number , 2011 = P2 text05
			query.setCriteria("[1001] starts with '" + number + "' or [2011] starts with '" + number + "'");
			ITable results = query.execute();
			if (results.size() > 0)
				return true;
			else
				return false;

		} catch (APIException apie) {
			apie.printStackTrace();
			logger.logException(apie);
			throw apie;
		} catch (Exception e) {
			e.printStackTrace();
			logger.logException(e);
			throw e;
		}
	}

	@Override
	public boolean setOutputValue(IRow row, IDataObject ido, String autoOutput, Ini ini, Log logger) {
		logger.log(">Reset Item Number...");
		try {
			ITable redLineTitleBlock = ido.getTable(ItemConstants.TABLE_REDLINETITLEBLOCK);
			IRow redlineRow = (IRow) redLineTitleBlock.iterator().next();
			ICell cell = redlineRow.getCell(ItemConstants.ATT_TITLE_BLOCK_NUMBER);
			cell.setValue(autoOutput.toUpperCase());
			IItem item = (IItem) ido;
			item.setValue(ItemConstants.ATT_PAGE_TWO_TEXT05, autoOutput.toUpperCase());
			failuresetname = false;
		} catch (Exception e) {
			failuresetname = true;
			errorCount++;
			failuremessage = failuremessage + " " + e.getMessage();
		}
		return failuresetname;
	}
}
