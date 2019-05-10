package common;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

import com.agile.api.CommonConstants;
import com.agile.api.IAgileList;
import com.agile.api.ICell;
import com.agile.api.IDataObject;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;

import Super.SN04_2_SetOutput;
import jxl.Cell;
import record.Log;

public class SN04_2_CommonUtil {
	
	public static ArrayList<Cell> checkMinorCategory(IDataObject ido,Map map,Log logger) throws Exception {
		
		ArrayList<Cell> cellList = null;
		ArrayList<Cell> tmp = null;
		for (Object key : map.keySet()) {
            String keyS = key.toString();
            // 表示無須根據任何分類判斷，直接回傳即可
            if("".equals(keyS)) {
            	cellList = (ArrayList<Cell>) map.get(key);
            	break;
            }
            String[] splitKey = keyS.split(",");
            if(splitKey.length<3) {
            	logger.log("Excel維護錯誤 - " + ido.getAgileClass().getAPIName() + "行,順序1");
            	return cellList;
            }
            ITable columnTable = null;
            if(splitKey[1].equalsIgnoreCase("P3"))
            	columnTable = ido.getTable(CommonConstants.TABLE_PAGETHREE);
            else if(splitKey[1].equalsIgnoreCase("P2"))
            	columnTable = ido.getTable(CommonConstants.TABLE_PAGETWO);
            else if(splitKey[1].equalsIgnoreCase("tb"))
            	columnTable = ido.getTable(ItemConstants.TABLE_TITLEBLOCK);
            Iterator it = columnTable.iterator();
            IRow row = (IRow) it.next();
            logger.log("Test:"+ido.getName());
            // 除判斷值以外內容都代入Others
            if("".equalsIgnoreCase(splitKey[2].trim())) {
            	logger.log("Excel維護錯誤 - " + ido.getAgileClass().getAPIName() + "行,順序1");
            	return cellList;
            }
            if("Others".equalsIgnoreCase(splitKey[2].trim())) {
            	tmp = (ArrayList<Cell>) map.get(key);
            	continue;
            }
            ICell cell = row.getCell(splitKey[0]);
            IAgileList list = (IAgileList) cell.getValue();
            IAgileList[] selected = list.getSelection();
            if(selected!=null) {
            	String targetValue = selected[0].getAPIName();
	            if(targetValue.trim().equals(splitKey[2].trim())) {
	            	logger.log(2,"取得小分類: "+selected[0].getValue());
	            	logger.log(2,"小分類 API Name: "+targetValue);
	            	cellList = (ArrayList<Cell>) map.get(key);
	            }
            }
            
        }
		if(cellList==null) {
			logger.log(2,"無對應小分類，代入預設分類");
			cellList = tmp;
		}
		return cellList;
	}
}
