/**
 * Created by wuchangming on 17/1/6.
 */

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 读取上篇中的xls文件的内容，并打印出来
 * @author Administrator
 *
 */
public class ReadExcelTable{



    //通过对单元格遍历的形式来获取信息 ，这里要判断单元格的类型才可以取出值
    public static List<Map> readTable() throws Exception {
        Map map = new HashMap<String, String>();
        List<Map> list = new ArrayList<Map>();
        map.put("308", "招商银行"); // 招行
        map.put("306", "广发银行"); // 广发
        map.put("303", "光大银行"); // 光大
        map.put("304", "华夏银行"); // 华夏
        map.put("105", "建设银行"); // 建行
        map.put("305", "民生银行"); // 民生
        map.put("103", "农业银行"); // 农行
        map.put("310", "浦发银行"); // 浦发
        map.put("309", "兴业银行"); // 兴业
        map.put("104", "中国银行"); // 中行
        map.put("302", "中信银行"); // 中信
        map.put("102", "工商银行"); // 工行
        map.put("301", "交通银行"); // 交行
        map.put("4105840", "平安银行"); // 平安
        map.put("403", "邮政储蓄银行"); // 邮储
        InputStream ips = new FileInputStream("/Users/wuchangming/Documents/123.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(ips);
        Sheet sheet1 = wb.getSheetAt(0);
        for (Row row : sheet1) {

            String name = row.getCell(0).getStringCellValue();
            String phoneNumber = row.getCell(2).getStringCellValue();
            String bank = (String) map.get(row.getCell(3).getStringCellValue());
            String account = row.getCell(4).getStringCellValue();
            String idCard = row.getCell(7).getStringCellValue();

            if("null".equals(bank)){
                continue;
            }
            Map dataMap = new HashMap();
            dataMap.put("name", name);
            dataMap.put("day", "4");
            dataMap.put("bank", bank);
            dataMap.put("bankAccount", account);
            dataMap.put("account", idCard);
            dataMap.put("phoneNUmber", phoneNumber);
            dataMap.put("shouquanren", name);
            dataMap.put("year", "2017");
            dataMap.put("month", "1");
            list.add(dataMap);

//            for (Cell cell : row) {
//                switch (cell.getCellType()) {
//                    case HSSFCell.CELL_TYPE_BOOLEAN:
//                        //得到Boolean对象的方法
//                        System.out.print(cell.getBooleanCellValue()+" ");
//                        break;
//                    case HSSFCell.CELL_TYPE_NUMERIC:
//                        //先看是否是日期格式
//                        if(HSSFDateUtil.isCellDateFormatted(cell)){
//                            //读取日期格式
//                            System.out.print(cell.getDateCellValue()+" ");
//                        }else{
//                            //读取数字
//                            System.out.print(cell.getNumericCellValue()+" ");
//                        }
//                        break;
//                    case HSSFCell.CELL_TYPE_FORMULA:
//                        //读取公式
//                        System.out.print(cell.getCellFormula()+" ");
//                        break;
//                    case HSSFCell.CELL_TYPE_STRING:
//                        //读取String
//                        System.out.print(cell.getRichStringCellValue().toString()+" ");
//                        break;
//                    default:
//                        System.out.println();
//                }
//
//            }
        }
        return list;

    }

//    //直接抽取excel中的数据
//    public static void extractor() throws Exception{
//        InputStream ips=new FileInputStream("d://test2-1.xls");
//        HSSFWorkbook wb=new HSSFWorkbook(ips);
//        ExcelExtractor ex=new ExcelExtractor(wb);
//        ex.setFormulasNotResults(true);
//        ex.setIncludeSheetNames(false);
//        String text=ex.getText();
//        System.out.println(text);
//    }


}
