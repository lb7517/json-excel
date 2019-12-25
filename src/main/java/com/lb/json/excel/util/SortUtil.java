package com.lb.json.excel.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

import java.util.Collections;
import java.util.Comparator;
import java.util.List;

public class SortUtil {

    /**
     * 排序JSONArray
     * @param jsonArray
     * return JSONArray
     * */
    public static JSONArray jsonArraySort(JSONArray jsonArray){
        List<JSONObject> list = JSONArray.parseArray(jsonArray.toJSONString(), JSONObject.class);
        Collections.sort(list, new Comparator<JSONObject>() {
            @Override
            public int compare(JSONObject o1, JSONObject o2) {
                int a = (int) o1.get("id");
                int b = (int) o2.get("id");
                if(a > b){  //升序排列，降序改成a<b
                    return 1;
                } else if(a == b){
                    return 0;
                } else {
                    return -1;
                }
            }
        });
        JSONArray jsonArray1 = JSONArray.parseArray(list.toString());
        return jsonArray1;
    }
}
