package com.macro.mall.tiny.modules.contract.service;

import com.macro.mall.tiny.common.api.CommonResult;
import com.macro.mall.tiny.modules.contract.model.TContract;
import com.baomidou.mybatisplus.extension.service.IService;
import org.springframework.web.multipart.MultipartFile;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author mx
 * @since 2021-06-10
 */
public interface TContractService extends IService<TContract> {
    CommonResult importExcel(MultipartFile excelFile);
}
