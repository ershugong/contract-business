package com.macro.mall.tiny.modules.contract.vo;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

@Data
public class ContractVO extends PageVO{
    @TableId(value = "id", type = IdType.AUTO)
    private Long id;

    @ApiModelProperty(value = "序号")
    private Long sort;

    @ApiModelProperty(value = "合同编号")
    private String contractNum;

    @ApiModelProperty(value = "合同名称")
    private String contractName;

    @ApiModelProperty(value = "合同分类")
    private String contractClassify;

    @ApiModelProperty(value = "对方单位名称")
    private String customerCompany;

    @ApiModelProperty(value = "我方名称")
    private String myCompany;

    @ApiModelProperty(value = "签名日期")
    private String startTime;

    @ApiModelProperty(value = "有效期间")
    private String valid;

    @ApiModelProperty(value = "到期日")
    private String endTime;

    @ApiModelProperty(value = "距离到期时间（超过当前时间以复数表示）")
    private String compareDay;

    @ApiModelProperty(value = "备注")
    private String remark;

    @ApiModelProperty(value = "是否合同原件 1是 0否")
    private String isOriginal;

    @ApiModelProperty(value = "跟进人")
    private String followMan;

    @ApiModelProperty(value = "合同文件路径")
    private String contractFile;
}
