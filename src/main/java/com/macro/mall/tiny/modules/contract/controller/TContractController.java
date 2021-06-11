package com.macro.mall.tiny.modules.contract.controller;

import com.macro.mall.tiny.common.api.CommonPage;
import com.macro.mall.tiny.common.api.CommonResult;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import com.macro.mall.tiny.modules.contract.service.TContractService;
import com.macro.mall.tiny.modules.contract.model.TContract;

import java.io.*;
import java.util.List;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;


/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author mx
 * @since 2021-06-10
 */
@RestController
@RequestMapping("/tContract")
@CrossOrigin
public class TContractController {

    @Autowired
    public TContractService tContractService;

    @Value("${uploadFilePath}")
    public String uploadFilePath;

    @RequestMapping(value = "/create", method = RequestMethod.POST)
    @ResponseBody
    public CommonResult create(@RequestBody TContract tContract) {
        boolean success = tContractService.save(tContract);
        if (success) {
            return CommonResult.success(null);
        }
        return CommonResult.failed();
    }

    @RequestMapping(value = "/update/{id}", method = RequestMethod.POST)
    @ResponseBody
    public CommonResult update(@PathVariable Long id, @RequestBody TContract tContract) {
        tContract.setId(id);
        boolean success = tContractService.updateById(tContract);
        if (success) {
            return CommonResult.success(null);
        }
        return CommonResult.failed();
    }

    @RequestMapping(value = "/delete/{id}", method = RequestMethod.POST)
    @ResponseBody
    public CommonResult delete(@PathVariable Long id) {
        boolean success = tContractService.removeById(id);
        if (success) {
            return CommonResult.success(null);
        }
        return CommonResult.failed();
    }

    @RequestMapping(value = "/delete", method = RequestMethod.POST)
    @ResponseBody
    public CommonResult deleteBatch(@RequestParam("ids") List<Long> ids) {
        boolean success = tContractService.removeByIds(ids);
        if (success) {
            return CommonResult.success(null);
        }
        return CommonResult.failed();
    }


    @RequestMapping(value = "/listAll", method = RequestMethod.POST)
    @ResponseBody
    public CommonResult<List<TContract>> listAll() {
        List<TContract> tContractList = tContractService.list();
        return CommonResult.success(tContractList);
    }

    @RequestMapping(value = "/listByPage", method = RequestMethod.POST)
    @ResponseBody
    public CommonResult listByPage(@RequestParam(value = "pageSize",defaultValue = "10") Integer pageSize,
                                   @RequestParam(value = "pageIndex",defaultValue = "1") Integer pageIndex,
                                   @RequestParam(value = "contractNum",defaultValue = "") String contractNum,
                                   @RequestParam(value = "contractName",defaultValue = "") String contractName,
                                   @RequestParam(value = "contractClassify",defaultValue = "") String contractClassify
    ) {
        Page<TContract> page = new Page<>(pageIndex,pageSize);
        QueryWrapper<TContract> queryWrapper = new QueryWrapper<>();
        if (!"".equals(contractNum)){
            queryWrapper.like("contract_num",contractNum);
        }
        if (!"".equals(contractName)){
            queryWrapper.like("contract_name",contractName);
        }
        if (!"".equals(contractClassify)){
            queryWrapper.like("contract_classify",contractClassify);
        }

        Page<TContract> pageResult = tContractService.page(page, queryWrapper);
        return CommonResult.success(CommonPage.restPage(pageResult));
    }

    @RequestMapping(value = "/importExcel", method = RequestMethod.POST)
    @ResponseBody
    public CommonResult importExcel(@RequestParam(value = "excelFile") MultipartFile excelFile) {
        return tContractService.importExcel(excelFile);
    }

    @RequestMapping("/savePdfForFile")
    public CommonResult savePdfForFile(@RequestParam(value = "file") MultipartFile file){
        BufferedInputStream in=null;
        BufferedOutputStream out=null;
        String totalName = "";
        try{
            String fileName = file.getOriginalFilename();
            InputStream is = file.getInputStream();
            in=new BufferedInputStream(is);
            totalName = uploadFilePath + "/" + fileName;
            out=new BufferedOutputStream(new FileOutputStream(totalName));
            int len=-1;
            byte[] b=new byte[1024];
            while((len=in.read(b))!=-1){
                out.write(b,0,len);
            }
        }catch (IOException e){
            CommonResult.failed("上传文件失败");
        }finally {
            try{
                if(in != null){
                    in.close();
                }
                if(out != null){
                    out.close();
                }
            }catch (IOException e){
                CommonResult.failed("关闭流资源失败");
            }
        }

        return CommonResult.success("/upload/" + totalName);
    }
}

