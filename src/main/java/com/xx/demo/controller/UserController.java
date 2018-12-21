package com.xx.demo.controller;


import com.xx.demo.config.excelUtils;
import com.xx.demo.entity.User;
import com.xx.demo.service.UserService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.List;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author admin
 * @since 2018-12-21
 */
@Controller
@RequestMapping("/user")
public class UserController {
    @Autowired
    UserService us;


    @RequestMapping(value = "/export")
    public void export(HttpServletRequest request, HttpServletResponse response) throws Exception {

        List<User> users = us.selectList(null);

        //excel标题
        String []title = {"id","username","password","sex","birthday","address"};
        //excel文件名
        String fileName = "用户信息表" + System.currentTimeMillis()+".xls";
        //sheet名
        String sheetName = "zz";

        //System.out.println("-----------------------------------------"+users.size());
        String[][] content = new String[users.size()][];
        for(int i=0;i<users.size();i++){

            content[i] = new String[title.length];
            User user = users.get(i);
            content[i][0] = user.getId().toString();
            content[i][1] = user.getUsername().toString();
            content[i][2] = user.getPassword().toString();
            content[i][3] = user.getSex().toString();
            content[i][4] = user.getBirthday().toString();
            content[i][5] = user.getAddress().toString();
        }


        //创建HSSFWorkbook
        XSSFWorkbook wb = excelUtils.getHSSFWorkbook(sheetName,title,content,null);

        //相应到客户端
        try {
            this.setResponseHeader(response, fileName);
            OutputStream os = response.getOutputStream();
            wb.write(os);
            os.flush();
            os.close();
        }catch (Exception e){
            e.printStackTrace();
        }

    }

    //发送响应流方法
    private void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            try {
                fileName = new String(fileName.getBytes(),"ISO8859-1");
            } catch (UnsupportedEncodingException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            response.setContentType("application/octet-stream;charset=ISO8859-1");
            response.setHeader("Content-Disposition", "attachment;filename="+ fileName);
            response.addHeader("Pargam", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    @RequestMapping(value = "fileUpload",method = RequestMethod.POST)
    public ModelAndView fileUpload(@RequestParam("file") MultipartFile file){

        ModelAndView mav = new ModelAndView();

        if(file.isEmpty()){
            mav.setViewName("false");
            return mav;
        }
        String fileName = file.getOriginalFilename();
        int size = (int) file.getSize();
        String path = "F:/test";
        File dest = new File(path +"/"+fileName);
        if(!dest.getParentFile().exists()){
            dest.getParentFile().mkdir();
        }
        try {
            file.transferTo(dest); //保存文件
            //System.out.println(dest.getAbsolutePath());
            String filepath = dest.getAbsolutePath();
            excelUtils e1 = new excelUtils();
            List<User> users =  e1.readXls(filepath);
            for(User user : users){
                System.out.println(user.toString());
            }
            mav.addObject("users",users);
            mav.setViewName("success");
            return mav;
        } catch (IOException e) {
            e.printStackTrace();
            mav.setViewName("false");
            return mav;
        }

    }
}

