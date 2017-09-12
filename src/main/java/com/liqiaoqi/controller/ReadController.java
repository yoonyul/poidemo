package com.liqiaoqi.controller;

import com.liqiaoqi.service.ReadExelService;
import com.liqiaoqi.service.TyphoonInfluenceService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * Created by LQQ on 2017/9/11 0011.
 */
@RestController
public class ReadController {
    @Autowired
    ReadExelService readExelService;
    @Autowired
    TyphoonInfluenceService typhoonInfluenceService;

    @RequestMapping("/read")
    public String read() throws  Exception{
        readExelService.readTyphoonInfo();
        return "success";
    }

    @RequestMapping("/typhoonInfluence")
    public String typhoonInfluence() throws Exception{
        typhoonInfluenceService.main();
        return "success";
    }
}
