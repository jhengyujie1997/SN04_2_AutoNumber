package Super;

import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IIncorporateFileEventInfo;

import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;
import jxl.Sheet;
import record.Log;
import common.GetInf;
import util.AUtil;
import util.Ini;

import java.io.IOException;
import java.security.KeyStore.LoadStoreParameter;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;

//完好的
public abstract class SetOutput {
	static String EXCEL_FILE;
	protected static int errorCount;
	protected IAgileSession admin;
	protected String FILEPATH;
	protected boolean failuresetname = false;// 若都沒失敗就是false
	protected String failuremessage = "";
	boolean failure = false;

	public boolean getfailuresetname() {
		// 若有跑錯(通常是配方) 此時會回傳true
		return failuresetname;
	}

	public String getfailuremessage() {
		// 取得(通常是配方)名字為何錯誤的訊息
		return failuremessage;
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

	public boolean statusInConfig(IChange change, IAgileSession admin, Ini ini, Log logger) {
		try {
			String currentStatus = change.getStatus().getAPIName().toString();
			final String[] availableStatus = ini.getValue("Workflow", "Status").split(",");
			for (String status : availableStatus) {
				if (status.replaceAll(" ", "").equals(currentStatus.replaceAll(" ", "")))
					return true;
			}
		} catch (APIException e) {
			logger.log(e);
			try {
				logger.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
			return false;
		}
		return false;
	}

	public void action(IChange change, IAgileSession admin, Ini ini, Log logger) throws Exception {
		try {
			change = (IChange) admin.getObject(IChange.OBJECT_TYPE, change.getName());
//        IChange changeOrder = (IChange) change;
			logger.log(1, "Get Change as Admin:" + change);
			ITable affectedTable = GetInf.getAffectedTable(change);
			Iterator<?> it = affectedTable.iterator();
			IRow row;
			IItem item;

			// loop through affected Items
			while (it.hasNext()) {
				// get affected item
				row = (IRow) it.next();
				// 跳過廠區有值
				if (!row.getValue(12208).toString().equals(""))
					continue;
				item = (IItem) row.getReferent();
				logger.log(1, "物件: " + item);

				// parse definition for excel class
				String autoOutput = getAutoOutput(item, admin, ini, logger);

				// Check Value
				if (autoOutput.equals("找不到規則")) {
					failuremessage += autoOutput + "\n\r";
					errorCount++;
					continue;
				}
				if (autoOutput.contains("欄位為空")) {
					failuremessage += autoOutput + "\n\r";
					errorCount++;
					continue;
				} 
				if(autoOutput.contains("請檢查 Excel")) {
					failuremessage += autoOutput + "\n\r";
					errorCount++;	
					continue;
				}
				if(autoOutput.contains("Excel維護錯誤")) {
					failuremessage += autoOutput + "\n\r";
					errorCount++;
					continue;
				}
				if(autoOutput.contains("需增加資料")) {
					failuremessage += autoOutput + "\n\r";
					errorCount++;
					continue;
				}
				if(autoOutput.contains("流水號已滿")) {
					failuremessage += autoOutput + "\n\r";
					errorCount++;
					continue;
				}
				
				// Set Value
				failure = setOutputValue(row, row.getReferent(), autoOutput, ini, logger);
				if (failure)
					continue;
			}
		} catch (APIException e) {
			e.printStackTrace();
			throw e;
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		} finally {
			try {
				logger.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	//
	public abstract String getAutoOutput(IItem item, IAgileSession session, Ini ini, Log logger) throws APIException, Exception;

	public abstract boolean setOutputValue(IRow row, IDataObject ido, String autoOutput, Ini ini, Log logger)
			throws APIException;

	protected int findClassRow(String agileAPIName, Sheet sheet, IItem item, Ini ini, Log logger) throws APIException {

		for (int i = 1; i <= sheet.getRows(); i++) {
			
			Cell cell1 = sheet.getCell(1, i);//Excel API
			if(cell1.getType() == CellType.EMPTY)continue;
			LabelCell labelCell = (LabelCell) cell1;
			String classAPIName = labelCell.getString();
			if(classAPIName.toLowerCase().equals("end"))break;
			Cell cell2 = sheet.getCell(2, i);//小分類
			if (classAPIName.toLowerCase().contains((agileAPIName.toLowerCase()))) {// Agile Class 要為APIname
				if (cell2.getType() != CellType.EMPTY) {
					LabelCell labelCell3 = (LabelCell) cell2;// 正式換成1
					if (item.getValue(ItemConstants.ATT_PAGE_THREE_LIST03) == null) {
						if (item.getValue(ItemConstants.ATT_PAGE_THREE_LIST03).toString().equals("")) {
							logger.log("小分類需填入資料!!!!");
							failuremessage += "小分類需填入資料!!!!";
							return -1;
						}
					}
					String list03 = item.getValue(ItemConstants.ATT_PAGE_THREE_LIST03).toString();
					if (list03.toString().contains("紙管")) {
						return i;
					}
					else {
						i =i+1;
						return i;
					}
				} //End if - no 小分類
				else {
					return i;
				}
			}
		}
		failuremessage += "找不到對應的規則!跳過...";
		return -1;
	}

}
// subclass name probably can't be found due to chinese character vs apiname