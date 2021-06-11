package com.macro.mall.tiny.modules.contract.service.impl;

import com.alibaba.fastjson.JSONObject;
import com.macro.mall.tiny.common.api.CommonResult;
import com.macro.mall.tiny.modules.contract.model.TContract;
import com.macro.mall.tiny.modules.contract.mapper.TContractMapper;
import com.macro.mall.tiny.modules.contract.service.TContractService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.macro.mall.tiny.modules.contract.util.ExcelUtil;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author mx
 * @since 2021-06-10
 */
@Service
public class TContractServiceImpl extends ServiceImpl<TContractMapper, TContract> implements TContractService {

    @Override
    public CommonResult importExcel(MultipartFile excelFile) {
        long startTime=System.currentTimeMillis();
        String originalFilename;
        if (excelFile == null) {
            return CommonResult.failed("上传文件为空");
        }
        try {
            originalFilename = excelFile.getOriginalFilename();
            if (!(originalFilename.endsWith(".xls") || originalFilename.endsWith(".xlsx"))) {
                return CommonResult.failed("文件格式格式有问题，请上传以xlxs或者xls文件格式的excel文档");
            }
            List<TContract> list = null;
            /*根据文件格式解析文件*/
            if (originalFilename.endsWith(".xls")) {
                list = ExcelUtil.readXLS(excelFile);
            } else if (originalFilename.endsWith(".xlsx")) {
                list = ExcelUtil.readXLSX(excelFile);
            }
            /*批量插入数据*/
            if(!CollectionUtils.isEmpty(list)){
                this.saveBatch(list);
            }
        } catch (Exception e) {
            e.printStackTrace();
            return CommonResult.failed("excel导入发生未知的错误，请稍后再试");
        }
        long hs=System.currentTimeMillis()-startTime;
        return CommonResult.success("excel导入成功，耗时"+hs+"毫秒！");
    }
}
