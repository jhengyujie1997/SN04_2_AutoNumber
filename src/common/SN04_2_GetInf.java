package common;

import com.agile.api.*;

import record.Log;

import java.util.HashMap;
import java.util.Iterator;

import javax.swing.plaf.multi.MultiLabelUI;

public class SN04_2_GetInf {
	public static String getCurrentUser(IAgileSession session) throws APIException {
		IUser currentUser = session.getCurrentUser();
		String userName = currentUser.getName();

		return userName;
	}

	/*
	 * 取得 change 中的 Table - AffectedItems
	 */
	public static ITable getAffectedTable(IChange change) throws APIException {
		ITable table;
		table = change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
		return table;
	}

	/*
	 * 根據 Table 的 APIName 與物件的 APIName ，取得需要的資訊並回傳
	 */
	public static String getObjectValue(IItem item, Log logger, IAgileSession admin, String[] configs)
			throws APIException {

		String Table = configs[0].replaceAll(" ", "");
		String Object_APIName = configs[1].replaceAll(" ", "");
		logger.log(1, "Config：" + Table + " / " + Object_APIName);
		String returnString = "";
		ITable table = item.getTable(Table);
		logger.log(1, "開始取得 Table：" + table.getName() + " 的資訊...");
		if (table.size() == 0) {
			returnString = item.getName() + " / Table " + table.getName() + " 需增加資料!! ";
			return returnString;
		} else {
			Iterator Iterator = table.iterator();
			while (Iterator.hasNext()) {
				IRow row = (IRow) Iterator.next();
				IDataObject ido = row.getReferent();
				// 製造商縮寫為製造商名稱
				String target = "";
				if (ido.getValue(Object_APIName) == null) {
					return Object_APIName + "欄位為空!!";
				} else {
					target = ido.getValue(Object_APIName).toString();
					return target;
				}
			}
		}
		return "";
	}

	public static void resetStatus(IChange change, IAgileSession session, Log logger) {
		IUser user = null;
		try {
			user = session.getCurrentUser();
		} catch (APIException e) {
			e.printStackTrace();
		}
		try {
			if (user.hasPrivilege(UserConstants.PRIV_CHANGESTATUS, change)) {
				IChange changeOrder = (IChange) session.getObject(IChange.OBJECT_TYPE, change.getName());
				IStatus currentStatus = changeOrder.getStatus();
				IWorkflow wf = changeOrder.getWorkflow();
				for (int i = 0; i < wf.getStates().length; i++) {
					if (currentStatus.equals(wf.getStates()[i])) {
						IStatus nextStatus = wf.getStates()[i - 1];
						changeOrder.changeStatus(nextStatus, false, "", false, false, null, null, null, false);
						return;
					}
				}
			}
		} catch (APIException e) {
			logger.log("退站出錯");
		}
	}

	/*
	 * Get agile list
	 */
	public static IAdminList getAgileList(IAgileSession session, String listName) throws APIException {
		IAdmin admin = session.getAdminInstance();
		IListLibrary listLib = admin.getListLibrary();
		IAdminList adminList = listLib.getAdminList(listName);
		return adminList;
	}

	public static String getcolunmValue(IItem item, Log logger, IAgileSession admin, String colunm_APIName, String page,
			String acceptNull) throws APIException {

		ICell cell = null;
		boolean priveledge=true;
		// 判斷 page2 page3
		if (page.charAt(1) == '3') {
			ITable pagethreeTable = item.getTable(ItemConstants.TABLE_PAGETHREE);
			Iterator it = pagethreeTable.iterator();
			IRow row = (IRow) it.next();
			cell = row.getCell(colunm_APIName);
		} else {
			ITable pagetwoTable = item.getTable(ItemConstants.TABLE_PAGETWO);
			Iterator it = pagetwoTable.iterator();
			IRow row = (IRow) it.next();
			cell = row.getCell(colunm_APIName);
		}
		
		//取得是空值
		if (cell == null)
			return item.getName() + " 中無 " + colunm_APIName + " ( 請檢查 Excel )... ";
		
		logger.log(1, "開始取得" + cell.getName() + " 的值...");
		logger.log(1, cell.getName() + " value：" + cell.getValue());
		
		//取得不是空值，但可為空值
		if (acceptNull != null && cell.getValue().toString()!=null) {
			return cell.getValue().toString();
		}
		
		//取得不是空值
		if (cell.getValue() != null) {
			//但是空字串
			if (cell.getValue().toString().equals("")) {
				return item.getName() + " 中 " + cell.getName() + " 欄位為空!!";
			}else {
				//不是空字串，所以回傳值
				return cell.getValue().toString();
			}	
			//取得是空值
		} else
			return item.getName() + " 中 " + cell.getName() + " 欄位為空!!";
	}

	public static String getListValue(IItem item, Log logger, IAgileSession admin, String colunm_APIName, String page,
			String check) throws APIException {

		ICell cell = null;
		// 判斷 page2 page3
		if (page.toUpperCase().equals("P3")) {
			ITable pagethreeTable = item.getTable(ItemConstants.TABLE_PAGETHREE);
			Iterator it = pagethreeTable.iterator();
			IRow row = (IRow) it.next();
			cell = row.getCell(colunm_APIName);
		} else {
			ITable pagetwoTable = item.getTable(ItemConstants.TABLE_PAGETWO);
			Iterator it = pagetwoTable.iterator();
			IRow row = (IRow) it.next();
			cell = row.getCell(colunm_APIName);
		}

		if (cell == null)
			return item.getName() + " 中無 " + colunm_APIName + " ( 請檢查 Excel ) ";

		logger.log(1, "開始取得" + cell.getName() + " 的清單...");
		IAgileList cl = (IAgileList) cell.getValue();
		String value = cell.getValue().toString();
		String value_Description = "";
		String returnValue = "";
		IAgileList[] selected = cl.getSelection();

		if (selected == null || selected.length == 0)
			return item.getName() + " 中 " + cell.getName() + "  欄位為空!! ( 請檢查 Excel )";
		if (selected != null && selected.length > 0) {

			if(selected[0].getDescription()==null)
				return item.getName() + " 中 " + cell.getName() + " list description 維護錯誤";
			if(selected[0].getDescription().toString().trim().equals(""))
				return item.getName() + " 中 " + cell.getName() + " list description 維護錯誤";
			if(selected[0].getDescription().toString().trim().equalsIgnoreCase("EMPTY"))
				return "";
			value_Description = selected[0].getDescription();
			String[] values = value_Description.split("@");

			if (check.equals("R")) {// Description 取得 @ 右邊值
				if (value_Description.indexOf("@") == value_Description.length() - 1)
					return item.getName() + " 中 " + cell.getName() + " list description  @右邊 欄位為空!!";
				returnValue = values[1].trim();/* .replaceAll(" ", ""); */
				if("EMPTY".equalsIgnoreCase(returnValue))
					return "";
				return returnValue;
			} else if (check.equals("L")) {// Number 取得 @ 左邊值
				returnValue = values[0].trim();/* .replaceAll(" ", ""); */
				if (returnValue.equals(""))
					return item.getName() + " 中 " + cell.getName() + " list description  左邊@ 欄位為空!!";
				else {
					if("EMPTY".equalsIgnoreCase(returnValue))
						return "";
					return returnValue;
				}
			}
		}
		return "";
	}

	public static boolean check(char c) {

		boolean check = true;
		switch (c) {
		case '2':
			check = false;
		case '3':
			check = false;
		case 'r':
			check = false;
		case 'l':
			check = false;
		case 'n':
			check = false;
		}

		return check;
	}
}
