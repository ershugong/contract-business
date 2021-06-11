package com.macro.mall.tiny.modules.contract.util;

import com.macro.mall.tiny.modules.contract.model.TContract;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelUtil {
    /**
     * 解析.xls格式的excel文件
     * @param file
     * @return
     * @throws IOException
     */
    public static List<TContract> readXLS(MultipartFile file) throws IOException {
        InputStream is = file.getInputStream();
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
        TContract entity;
        List<TContract> list = new ArrayList<>();
        // 循环工作表Sheet
        // 循环行Row
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
        HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
        for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
            HSSFRow hssfRow = hssfSheet.getRow(rowNum);
            if (hssfRow != null) {
                entity = new TContract();
                entity.setSort(new Double(hssfRow.getCell(0).getNumericCellValue()).longValue());
                entity.setContractNum(hssfRow.getCell(1).getStringCellValue());
                entity.setContractName(hssfRow.getCell(2).getStringCellValue());
                entity.setContractClassify(hssfRow.getCell(3).getStringCellValue());
                entity.setCustomerCompany(hssfRow.getCell(4).getStringCellValue());
                entity.setMyCompany(hssfRow.getCell(5).getStringCellValue());

                HSSFCell startTimeCell = hssfRow.getCell(6);
                entity.setStartTime(sdf.format(startTimeCell.getDateCellValue()));

                entity.setValid(hssfRow.getCell(7).getStringCellValue());

                HSSFCell endTimeCell = hssfRow.getCell(8);
                entity.setEndTime(Cell.CELL_TYPE_STRING == endTimeCell.getCellType() ? endTimeCell.getStringCellValue() : sdf.format(endTimeCell.getDateCellValue()));

                HSSFCell companyCell = hssfRow.getCell(9);
                companyCell.setCellType(Cell.CELL_TYPE_STRING);
                entity.setCompareDay(companyCell.getStringCellValue());

                HSSFCell remarkCell = hssfRow.getCell(10);
                entity.setRemark(remarkCell == null ? "" : remarkCell.getStringCellValue());

                entity.setIsOriginal(hssfRow.getCell(11).getStringCellValue());
                entity.setFollowMan(hssfRow.getCell(12).getStringCellValue());
                //excel导入默认采用合同编号，作为附件的文件名
                entity.setContractFile("/upload/"+entity.getContractNum()+".pdf");
                list.add(entity);
            }
        }

        if(is!=null){
            is.close();
        }
        return list;
    }

    /**
     * 解析.xlsx格式的excel文件
     * @param file
     * @return
     * @throws IOException
     */
    public static List<TContract> readXLSX(MultipartFile file) throws IOException {
        InputStream is = file.getInputStream();
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
        TContract entity;
        List<TContract> list = new ArrayList<>();
        // 循环工作表Sheet
        // 循环行Row
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
        for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
            XSSFRow xssfRow = xssfSheet.getRow(rowNum);
            if (xssfRow != null) {
                entity = new TContract();
                entity.setSort(new Double(xssfRow.getCell(0).getNumericCellValue()).longValue());
                entity.setContractNum(xssfRow.getCell(1).getStringCellValue());
                entity.setContractName(xssfRow.getCell(2).getStringCellValue());
                entity.setContractClassify(xssfRow.getCell(3).getStringCellValue());
                entity.setCustomerCompany(xssfRow.getCell(4).getStringCellValue());
                entity.setMyCompany(xssfRow.getCell(5).getStringCellValue());

                XSSFCell startTimeCell = xssfRow.getCell(6);
                entity.setStartTime(sdf.format(startTimeCell.getDateCellValue()));

                entity.setValid(xssfRow.getCell(7).getStringCellValue());

                XSSFCell endTimeCell = xssfRow.getCell(8);
                entity.setEndTime(Cell.CELL_TYPE_STRING == endTimeCell.getCellType() ? endTimeCell.getStringCellValue() : sdf.format(endTimeCell.getDateCellValue()));

                XSSFCell companyCell = xssfRow.getCell(9);
                companyCell.setCellType(Cell.CELL_TYPE_STRING);
                entity.setCompareDay(companyCell.getStringCellValue());

                XSSFCell remarkCell = xssfRow.getCell(10);
                entity.setRemark(remarkCell == null ? "" : remarkCell.getStringCellValue());

                entity.setIsOriginal(xssfRow.getCell(11).getStringCellValue());
                entity.setFollowMan(xssfRow.getCell(12).getStringCellValue());
                //excel导入默认采用合同编号，作为附件的文件名
                entity.setContractFile("/upload/"+entity.getContractNum()+".pdf");
                list.add(entity);
            }
        }

        if(null!=is){
            is.close();
        }
        return list;
    }
}
