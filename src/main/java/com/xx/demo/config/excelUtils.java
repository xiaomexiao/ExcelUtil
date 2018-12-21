package com.xx.demo.config;

import com.xx.demo.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class excelUtils {

    //excel表导出
    public static XSSFWorkbook getHSSFWorkbook(String sheetName,String[] title,String[][] values,XSSFWorkbook wb){
        //第一步，创建一个HSSFWorkbook，对应一个Excel文件
        if (wb == null){
            wb = new XSSFWorkbook();
        }

        //第二步，在workbook中添加一个sheet，对应excel文件中的sheet
        XSSFSheet sheet = wb.createSheet(sheetName);

        //第三步，在sheet中添加表头第0行，注意老版本poi对excel的行数列数有限制
        XSSFRow row = sheet.createRow(0);

        //第四步，创建单元格,并设置表头   设置表头居中
        XSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);//创建一个居中样式

        //声明列对象
        XSSFCell cell = null;

        //创建标题
        for (int i=0;i<title.length;i++){
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        //创建内容
        for(int i=0;i<values.length;i++){
            row = sheet.createRow(i + 1);
            for(int j=0;j<values[i].length;j++){
                //将内容按顺序赋给对应的列对象
                row.createCell(j).setCellValue(values[i][j]);
            }
        }
        return wb;
    }


    //excel表导入
    public List<User> readXls(String path) throws IOException {
        List<User> users = new ArrayList<User>();
        Workbook workbook = new XSSFWorkbook(new FileInputStream(path));
        //根据名称获取指定sheet对象
        Sheet hssfSheet = workbook.getSheet("zz");
        for (Row row : hssfSheet) {
            int rowNum = row.getRowNum();
            //从第二行开始读数据了
            if (rowNum == 0) {
                continue;
            }

            for (Cell cell : row) {
                cell.setCellType(CellType.STRING);
            }

            String sid = row.getCell(0).getStringCellValue();
            int id = Integer.parseInt(sid);
            String username = row.getCell(1).getStringCellValue();
            String password = row.getCell(2).getStringCellValue();
            String ssex = row.getCell(3).getStringCellValue();
            int sex = Integer.parseInt(ssex);

            //excel中大的数字和date类型的日期会被cell的科学记数法解析成什么奇怪的东西
            String sbirthday = row.getCell(4).getStringCellValue();
//                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-mm-dd");
//                Date birthday = null;
//                try {
//                    birthday = sdf.parse(sbirthday);
//                } catch (ParseException e) {
//                    e.printStackTrace();
//                }
            //Date birthday = SimpleDateFormat.
            String address = row.getCell(5).getStringCellValue();
            User user = new User(id, username, password, sex, sbirthday, address);
            users.add(user);

        }

        return users;
    }
}
