package com.lb.json.excel;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.lb.json.excel.util.JsonExcelUtil;
import com.lb.json.excel.util.JsonUtil;
import com.lb.json.excel.util.SortUtil;

public class JsonExcel {

    public static void main(String args[]){
        JsonExcel.jsonToExcel();  // json转excel
        JsonExcel.excelToJson();  // excel转json
    }

    private static void jsonToExcel(){
        // excel路径
        String path = "E:\\5g\\文档\\mec文档\\模板配置信息\\excel\\漏洞模板.xlsx";
        // json路径
        String path1 = "E:\\5g\\文档\\mec文档\\模板配置信息\\漏洞模板.json";
        // excel表题头
        String[] headers = {"id", "name", "description"};
        // 转换成json字符串
        String a = JsonUtil.readJsonFile(path1);
        JSONObject jsonObject = JSONObject.parseObject(a);
        JSONArray jsonArray1 = (JSONArray) jsonObject.get("data");
        // 排序
        JSONArray jsonArray = SortUtil.jsonArraySort(jsonArray1);
        JsonExcelUtil.jsonToExcel(jsonArray, path, headers);  // excel转json
    }

    private static void excelToJson(){
        // excel路径
        String path = "E:\\5g\\文档\\mec文档\\模板配置信息\\excel\\漏洞模板.xlsx";
        JsonExcelUtil.excelToJson(path);  // excel转json
    }
}
