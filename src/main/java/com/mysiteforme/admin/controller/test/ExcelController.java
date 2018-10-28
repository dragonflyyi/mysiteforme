package com.mysiteforme.admin.controller.test;

import com.aliyun.oss.HttpMethod;
import com.baomidou.mybatisplus.mapper.EntityWrapper;
import com.mysiteforme.admin.base.BaseController;
import com.mysiteforme.admin.entity.Log;
import com.mysiteforme.admin.service.LogService;
import com.mysiteforme.admin.service.UserService;
import com.mysiteforme.admin.util.ExcelException;
import com.mysiteforme.admin.util.ExcelTool;
import com.xiaoleilu.hutool.poi.excel.ExcelUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by dragonfly on 2018/10/22.
 */

@RestController
@RequestMapping(path = "/test/excel")
public class ExcelController extends BaseController{

    @Autowired
    private LogService logService;

    @RequestMapping(path = "/my")
    public String  exportFile(String type) throws ExcelException {

        EntityWrapper<Log>  wrap = new EntityWrapper();
        wrap.eq("type", type);
        List<Log> resultList = logService.selectList(wrap);



        String exportFileName ="/Users/dragonfly/test2.xls";
        ExcelTool.list2File(resultList, Log.class, exportFileName);

        return "OK";

    }
}
