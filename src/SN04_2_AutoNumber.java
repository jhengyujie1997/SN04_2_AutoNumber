
import java.io.Closeable;
import java.io.IOException;
import java.sql.Connection;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;

import javax.swing.plaf.basic.BasicInternalFrameTitlePane.IconifyAction;

import com.agile.api.APIException;
import com.agile.api.ChangeConstants;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;
import com.agile.px.IWFChangeStatusEventInfo;

import AutoOutput.SN04_2_SetNumber;
import record.Log;
import util.AUtil;
import util.Ini;

public class SN04_2_AutoNumber implements ICustomAction, IEventAction {

	protected Ini ini;
	protected Log log;
	protected IAgileSession admin;
	protected String FILEPATH;
	SN04_2_SetNumber number;

	/**********************************************
	 * 初始化
	 **********************************************/
	private void init() {
		ini = new Ini("D:\\Anselm_Program_Data\\Config.ini");
		log = new Log("SN04_2_AutoNumber");
		FILEPATH = ini.getValue("Program Use", "LogFile") + "SN04_2_AutoNumber_";
		admin = AUtil.getAgileSession(ini, "AgileAP");
	}

	/**********************************************
	 * 關閉建構
	 **********************************************/
	public void close() {
		ini = null;
		try {
			log.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**********************************************
	 * 主程式執行 - PX
	 **********************************************/
	@Override
	public ActionResult doAction(IAgileSession iAgileSession, INode iNode, IDataObject change) {
		init();
		log.logSeparatorBar();
		try {
			log.setLogFile(FILEPATH+change+".log");
		} catch (IOException e) {
			e.printStackTrace();
		}
		String result = "";
		int errorCount = 0;
		IChange changeOrder = null;
		try {
			changeOrder = (IChange) change;
		} catch (Exception e) {
			log.log(e);
			result = "Object is not Change!";
			close();
			return new ActionResult(ActionResult.STRING, result);
		}
		SN04_2_SetNumber number = new SN04_2_SetNumber();
		// Auto Number
		try {
			number.action(changeOrder, admin, ini, log);
		} catch (Exception e) {
			close();
			return new ActionResult(ActionResult.EXCEPTION, e);
		}
		errorCount = number.getErrorCount();
		result = errorCount == 0 ? "程式執行完成" : "執行自動編碼總共有" + errorCount + "筆條件失敗，請檢查log檔" + "\n\r";
		if (errorCount != 0) {
			result += number.getfailuremessage();
			number.resetCount();
			number.resetfailuremessage();
			return new ActionResult(ActionResult.STRING, result);
		}

		close();
		if("".equals(result))
			result = "Done";
		return new ActionResult(ActionResult.STRING, result);
	}

	/**********************************************
	 * 主程式執行 - Event
	 **********************************************/
	@Override
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo req) {
		init();
		log.logSeparatorBar();
		IWFChangeStatusEventInfo info = (IWFChangeStatusEventInfo) req;
		String result = "";
		int errorCount = 0;
		try {
			log.setLogFile(FILEPATH);
		} catch (IOException e) {
			e.printStackTrace();
		}
		log.logSeparatorBar();
		IChange changeOrder = null;
		try {
			changeOrder = (IChange) info.getDataObject();
		} catch (APIException e) {
			log.log(e);
			result = "Object is not Change !!!";
			close();
			return new EventActionResult(req, new ActionResult(ActionResult.STRING, result));
		}

		SN04_2_SetNumber number = new SN04_2_SetNumber();

		// Auto Number
		try {
			number.action(changeOrder, admin, ini, log);
		} catch (Exception e) {
			close();
			return new EventActionResult(req, new ActionResult(ActionResult.EXCEPTION, e));
		}
		errorCount = number.getErrorCount();
		result = errorCount == 0 ? "程式執行完成" : "執行自動編碼總共有" + errorCount + "筆條件失敗，請檢查log檔" + "\n\r";
		if (errorCount != 0) {
			result += number.getfailuremessage();
			number.resetCount();
			return new EventActionResult(req, new ActionResult(ActionResult.STRING, result));
		}

		if (errorCount == 0)
			result = "End#";

		close();
		return new EventActionResult(req, new ActionResult(ActionResult.STRING, result));
	}
}
