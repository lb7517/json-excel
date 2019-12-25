package com.lb.json.excel.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class JsonExcelUtil {
    /**
     * json转换成excel
     * @param jsonArray
     * @param excelPath   excel路径
     * @param headers     excel表题头 如: {"id", "name"}
     * */
    public static void jsonToExcel(JSONArray jsonArray, String excelPath, String[] headers){
        Sheet sheet;
        try{
            Workbook wb = new XSSFWorkbook();
            sheet = wb.createSheet("sheet1");
            Row row1 = sheet.createRow(0);
            for(int i = 0; i< headers.length; i++){//表头赋值
                row1.createCell(i).setCellValue(headers[i]);
            }
            for(int i = 0; i< headers.length; i++){//表头赋值
                row1.createCell(i).setCellValue(headers[i]);
            }
            if (jsonArray.size() > 0) {//单元格赋值
                for (int i = 0; i < jsonArray.size(); i++) {
                    Row row2 = sheet.createRow(i + 1);
                    JSONObject json = jsonArray.getJSONObject(i); // 遍历 jsonarray
                    for(int j = 0;j < headers.length; j++){
                        row2.createCell(j).setCellValue(json.get(headers[j]).toString());//赋值
                    }
                }
            }

            FileOutputStream fos = new FileOutputStream(excelPath);
            wb.write(fos);
            fos.close();
        }catch (Exception e){
            System.out.println("e: "+e);
        }
    }

    /**
     * excel转换成json
     * @param excelPath   excel路径
     *  return JSONArray
     * */
    public static JSONArray excelToJson(String excelPath){
        Sheet sheet;
        try{
            OPCPackage pkg = OPCPackage.open(excelPath);
            Workbook wb = new XSSFWorkbook(pkg);
            // 获取第一个sheet
            sheet = wb.getSheetAt(0);
            JSONArray jsonArray = new JSONArray();
            // 获取第一行,即标题
            Row row1 = sheet.getRow(0);
            // 获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            // 获取最大列数
            int column = row1.getPhysicalNumberOfCells();
            for(int i = 1; i< rownum; i++){//表头赋值
                JSONObject map = new JSONObject();
                Row row2 = sheet.getRow(i);
                for(int j = 0; j < column; j++){
                    // cell表示格子
                    Cell cell = row2.getCell(j);
                    // 注意区分类型，否者有些显示会异常(如为空)
                    switch (cell.getCellType()){
                        // 注意此处不能加类名,即不能写成CellType.NUMERIC
                        case NUMERIC: {
                            map.put(row1.getCell(j).getStringCellValue(), cell.getNumericCellValue());
                            break;
                        }
                        case FORMULA:{
                            //判断cell是否为日期格式
                            if (DateUtil.isCellDateFormatted(cell)) {
                                //转换为日期格式YYYY-mm-dd
                                map.put(row1.getCell(j).getStringCellValue(), cell.getDateCellValue());
                            } else {
                                //数字
                                map.put(row1.getCell(j).getStringCellValue(), cell.getNumericCellValue());
                            }
                            break;
                        }
                        case STRING: {
                            map.put(row1.getCell(j).getStringCellValue(), cell.getStringCellValue());
                            break;
                        }
                        default:
                            map.put(row1.getCell(j).getStringCellValue(), "");
                    }
                }
                jsonArray.add(map);
            }
            System.out.println("jsonArray:"+jsonArray);
            return jsonArray;
        }catch (Exception e){
            System.out.println("e: "+e);
            JSONArray jsonArray = new JSONArray();
            return jsonArray;
        }
    }
}
