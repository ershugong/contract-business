package com.macro.mall.tiny.modules.contract.service.impl;

import cn.hutool.core.collection.CollectionUtil;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import com.macro.mall.tiny.common.api.CommonResult;
import com.macro.mall.tiny.modules.contract.model.TContract;
import com.macro.mall.tiny.modules.contract.mapper.TContractMapper;
import com.macro.mall.tiny.modules.contract.service.TContractService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.macro.mall.tiny.modules.contract.util.ExcelUtil;
import com.macro.mall.tiny.modules.contract.vo.ContractVO;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.Date;
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

    @Value("${uploadFilePath}")
    public String uploadFilePath;

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

    @Override
    public CommonResult exportExcel(ContractVO contractVO) {
        Page<TContract> page = new Page<>(contractVO.getPageIndex(),contractVO.getPageSize());
        QueryWrapper<TContract> queryWrapper = new QueryWrapper<>();
        if (!"".equals(contractVO.getContractNum())){
            queryWrapper.like("contract_num",contractVO.getContractNum());
        }
        if (!"".equals(contractVO.getContractName())){
            queryWrapper.like("contract_name",contractVO.getContractName());
        }
        if (!"".equals(contractVO.getContractClassify())){
            queryWrapper.like("contract_classify",contractVO.getContractClassify());
        }

        Page<TContract> pageResult = this.page(page, queryWrapper);
        List<TContract> list = pageResult.getRecords();
        if(CollectionUtil.isEmpty(list)){
            return CommonResult.failed("没有查询到相关的数据");
        }
        HSSFWorkbook wb = new HSSFWorkbook();//创建工作薄
        HSSFSheet sheet = wb.createSheet("合同列表");

        // 第四步，创建表头样式
        CellStyle headStyle = setHeadStyle(wb);
        // 给单元格内容设置另一个样式
        CellStyle cellStyle = setCellStyle(wb);
        int rowIndex = 0;
        int cellIndex = 0;
        Row titleRow = sheet.createRow(rowIndex);
        String [] titleText = {"序号","合同编号","合同名称","合同分类","对方单位名称","我方单位名称","签订日期","有效期间","到期日","是否到期","备注","合同原件","跟进人","合同附件地址"};
        for(int i=0;i<titleText.length;i++){
            Cell cell = titleRow.createCell(cellIndex++);
            cell.setCellStyle(headStyle);
            cell.setCellValue(titleText[i]);
        }
        //设置自动列宽(必须在单元格设值以后进行)
//        sheet.autoSizeColumn(0);
//        sheet.setColumnWidth(0, sheet.getColumnWidth(0) * 17 / 10);

        rowIndex++;
//        StringBuffer stringBuffer;
        int index = 1;
        DecimalFormat df=new DecimalFormat("0.00");
        for(TContract contract : list){
            cellIndex = 0;
            Row row = sheet.createRow(rowIndex++);
            Cell sortCell = row.createCell(cellIndex++);
            sortCell.setCellStyle(cellStyle);
            sortCell.setCellValue(contract.getSort());

            Cell contractNumCell = row.createCell(cellIndex++);
            contractNumCell.setCellStyle(cellStyle);
            contractNumCell.setCellValue(contract.getContractNum());

            Cell contractNameCell = row.createCell(cellIndex++);
            contractNameCell.setCellStyle(cellStyle);
            contractNameCell.setCellValue(contract.getContractName());

            Cell contractClassifyCell = row.createCell(cellIndex++);
            contractClassifyCell.setCellStyle(cellStyle);
            contractClassifyCell.setCellValue(contract.getContractClassify());

            Cell customerCompanyCell = row.createCell(cellIndex++);
            customerCompanyCell.setCellStyle(cellStyle);
            customerCompanyCell.setCellValue(contract.getCustomerCompany());

            Cell myCompanyCell = row.createCell(cellIndex++);
            myCompanyCell.setCellStyle(cellStyle);
            myCompanyCell.setCellValue(contract.getMyCompany());

            Cell startTimeCell = row.createCell(cellIndex++);
            startTimeCell.setCellStyle(cellStyle);
            startTimeCell.setCellValue(contract.getStartTime());

            Cell validCell = row.createCell(cellIndex++);
            validCell.setCellStyle(cellStyle);
            validCell.setCellValue(contract.getValid());

            Cell endTimeCell = row.createCell(cellIndex++);
            endTimeCell.setCellStyle(cellStyle);
            endTimeCell.setCellValue(contract.getEndTime());

            Cell compareDayCell = row.createCell(cellIndex++);
            compareDayCell.setCellStyle(cellStyle);
            compareDayCell.setCellValue(contract.getCompareDay());

            Cell remarkCell = row.createCell(cellIndex++);
            remarkCell.setCellStyle(cellStyle);
            remarkCell.setCellValue(contract.getRemark());

            Cell isOriginalCell = row.createCell(cellIndex++);
            isOriginalCell.setCellStyle(cellStyle);
            isOriginalCell.setCellValue(contract.getIsOriginal());

            Cell followManCell = row.createCell(cellIndex++);
            followManCell.setCellStyle(cellStyle);
            followManCell.setCellValue(contract.getFollowMan());

            Cell contractFileCell = row.createCell(cellIndex++);
            contractFileCell.setCellStyle(cellStyle);
            contractFileCell.setCellValue("http://123.207.88.92/contractApi" + contract.getContractFile());
            index++;

        }

        // 必须在单元格设值以后进行
        // 设置为根据内容自动调整列宽
        for (int k = 0; k < titleText.length; k++) {
            sheet.autoSizeColumn(k);
        }
        // 处理中文不能自动调整列宽的问题
        this.setSizeColumn(sheet, titleText.length);

        String result = downloadExcel(wb);
        if (result != null){
            return CommonResult.success("/upload/" + result + ".xls");
        } else {
            return CommonResult.failed("导出失败");
        }
    }

    private static HSSFCellStyle setHeadStyle(HSSFWorkbook wb) {
        HSSFCellStyle headStyle = wb.createCellStyle();
        // 设置背景颜色白色  HSSFColor.LIGHT_GREEN.index HSSFColor.GREY_25_PERCENT.index
        headStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        headStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        headStyle.setBorderBottom(CellStyle.BORDER_THIN);
        headStyle.setBottomBorderColor(HSSFColor.BLACK.index);
        headStyle.setBorderTop(CellStyle.BORDER_THIN);
        headStyle.setTopBorderColor(HSSFColor.BLACK.index);
        headStyle.setBorderLeft(CellStyle.BORDER_THIN);
        headStyle.setLeftBorderColor(HSSFColor.BLACK.index);
        headStyle.setBorderRight(CellStyle.BORDER_THIN);
        headStyle.setRightBorderColor(HSSFColor.BLACK.index);

        // 设置标题字体
        HSSFFont headFont = wb.createFont();
        // 设置字体大小
        headFont.setFontHeightInPoints((short) 12);
        // 设置字体
        headFont.setFontName("宋体");
        // 设置字体粗体
        headFont.setBold(true);
        // 把字体应用到当前的样式
        headStyle.setFont(headFont);
        return headStyle;
    }

    // 自适应宽度(中文支持)
    private void setSizeColumn(HSSFSheet sheet, int size) {
        for (int columnNum = 0; columnNum < size; columnNum++) {
            if(columnNum == 4){
                continue;
            }
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                HSSFRow currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }

                if (currentRow.getCell(columnNum) != null) {
                    HSSFCell currentCell = currentRow.getCell(columnNum);
                    if (currentCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
        }
    }

    /**
     * 设置单元格内容样式
     * @param wb
     * @return
     */
    private static HSSFCellStyle setCellStyle(HSSFWorkbook wb) {
        HSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setWrapText(true);
        // 设置标题字体
        HSSFFont cellFont = wb.createFont();
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBottomBorderColor(HSSFColor.BLACK.index);
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setTopBorderColor(HSSFColor.BLACK.index);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setLeftBorderColor(HSSFColor.BLACK.index);
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        cellStyle.setRightBorderColor(HSSFColor.BLACK.index);
        // 设置字体大小
//        cellFont.setFontHeightInPoints((short) 11);
        // 设置字体
//        cellFont.setFontName("等线");
        // 设置字体
        cellFont.setFontName("宋体");
        // 把字体应用到当前的样式
//        cellStyle.setFont(cellFont);
        return cellStyle;
    }

    public String downloadExcel(HSSFWorkbook wb) {
//        boolean flag = true;
//        Date date = new Date();
        String fileName = new Long(new Date().getTime()).toString();
        String filePath = uploadFilePath+ "/" + fileName + ".xls";
        File file = new File(filePath);

        FileOutputStream fos = null;
        try {

            fos = new FileOutputStream(file);
            wb.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                }catch (IOException e) {
                    e.printStackTrace();
                    return null;
                }
            }
        }
        return fileName;
    }
}
