package com.yinglian.modules.api.controller;

import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.shiro.authz.annotation.RequiresPermissions;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.alibaba.fastjson.JSONArray;
import com.yinglian.common.utils.SmsConstant;
import com.yinglian.modules.sys.service.SysConfigService;
import com.yinglian.utils.ShiroUtils;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.http.HttpUtil;
import cn.hutool.json.JSONObject;
import cn.hutool.json.JSONUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

/**
 * 贷后管理
 *
 * @author huyc
 * @date 2020/4/26 14:35:19
 */
@RestController
@RequestMapping("api/loanAfterManage")
public class LoanAfterManageController {

    @Value("${finance.service.api.host}")
    private String apiHost;

    @Value("${finance.service.api.after-manage.repayment-list}")
    private String repaymentList;

    @Value("${finance.service.api.after-manage.repayment-detail}")
    private String repaymentDetail;

    @Value("${finance.service.api.after-manage.operate-list}")
    private String operateList;

    @Value("${finance.service.api.after-manage.operate-detail}")
    private String operateDetail;

    @Value("${finance.service.api.after-manage.account-balance-warn}")
    private String accountBalanceWarn;

    @Value("${finance.service.api.after-manage.query-account-balance}")
    private String queryAccountBalance;

    @Value("${finance.service.api.after-manage.query-pay-off}")
    private String queryPayOff;

    @Autowired
    private SysConfigService configService;


    /**
     * 还款管理列表
     *
     * @param incomeNo    进件编号
     * @param loanNo      贷款编号
     * @param companyName 企业名称
     * @param isPayoff    是否还清
     * @param startDate   更新开始时间
     * @param endDate     更新结束时间
     * @return
     */
    @RequestMapping("/repaymentList")
    @RequiresPermissions("api:loanAfterManage:repaymentList")
    public String repaymentList(@RequestParam(required = false) String incomeNo,
                                @RequestParam(required = false) String loanNo,
                                @RequestParam(required = false) String companyName,
                                @RequestParam(required = false) Integer isPayoff,
                                @RequestParam(required = false) String startDate,
                                @RequestParam(required = false) String endDate,
                                @RequestParam Integer page,
                                @RequestParam Integer limit) {
        Map<String, Object> params = new HashMap<>(6);
        params.put("incomeNo", incomeNo);
        params.put("loanNo", loanNo);
        params.put("companyName", companyName);
        params.put("isPayoff", isPayoff);
        params.put("startDate", startDate);
        params.put("endDate", endDate);
        params.put("page", page.toString());
        params.put("limit", limit.toString());
        return HttpUtil.get(apiHost + repaymentList, params);

    }

    /**
     * 还款管理列表详情
     *
     * @param loanNo 借据编号
     * @return
     */
    @RequestMapping("/repaymentDetail")
    @RequiresPermissions("api:loanAfterManage:repaymentDetail")
    public String repaymentDetail(@RequestParam String loanNo) {
        Map<String, Object> params = new HashMap<>(1);
        params.put("loanNo", loanNo);
        return HttpUtil.get(apiHost + repaymentDetail, params);
    }

    /**
     * 查询经营数据列表
     *
     * @param incomeNo    进件编号
     * @param hotelName   酒店名称
     * @param companyName 企业名称
     * @param isPayoff    是否还清
     * @param type        酒店类型
     * @param startDate   创建开始时间
     * @param endDate     创建结束时间
     * @return
     */
    @RequestMapping("/operateList")
    @RequiresPermissions("api:loanAfterManage:operateList")
    public String operateList(@RequestParam(required = false) String incomeNo,
                              @RequestParam(required = false) String hotelName,
                              @RequestParam(required = false) String companyName,
                              @RequestParam(required = false) Integer isPayoff,
                              @RequestParam(required = false) Integer type,
                              @RequestParam(required = false) String startDate,
                              @RequestParam(required = false) String endDate,
                              @RequestParam Integer page,
                              @RequestParam Integer limit) {
        Map<String, Object> params = new HashMap<>(9);
        params.put("incomeNo", incomeNo);
        params.put("hotelName", hotelName);
        params.put("companyName", companyName);
        params.put("isPayoff", isPayoff);
        params.put("type", type);
        params.put("startDate", startDate);
        params.put("endDate", endDate);
        params.put("page", page.toString());
        params.put("limit", limit.toString());
        return HttpUtil.get(apiHost + operateList, params);
    }

    /**
     * 查询经营数据详情
     *
     * @param incomeNo          进件编号
     * @param companyCreditCode 企业信用代码
     * @param hotelType         酒店类型
     * @return
     */
    @RequestMapping("/operateDetail")
    @RequiresPermissions("api:loanAfterManage:operateDetail")
    public String operateDetail(@RequestParam String incomeNo,
                                @RequestParam String companyCreditCode,
                                @RequestParam Integer hotelType) {
        Map<String, Object> params = new HashMap<>(3);
        params.put("incomeNo", incomeNo);
        params.put("companyCreditCode", companyCreditCode);
        params.put("hotelType", hotelType.toString());
        return HttpUtil.get(apiHost + operateDetail, params);
    }

    /**
     * 导出经营数据详情
     *
     * @param incomeNo          进件编号
     * @param companyCreditCode 企业信用代码
     * @return
     */
    @RequestMapping("/exportOperateDetail")
    @SuppressWarnings("unchecked")
    @RequiresPermissions("api:loanAfterManage:exportOperateDetail")
    public void exportOperateDetail(HttpServletRequest request, HttpServletResponse response,
                                    @RequestParam String incomeNo,
                                    @RequestParam String companyCreditCode,
                                    @RequestParam String hotelType) throws Exception {
        Map<String, Object> params = new HashMap<>(2);
        params.put("incomeNo", incomeNo);
        params.put("companyCreditCode", companyCreditCode);
        params.put("hotelType", hotelType);

        String result = HttpUtil.get(apiHost + operateDetail, params);
        JSONObject json = JSONUtil.parseObj(result);
        Map<String, Object> map = JSONArray.parseObject(json.getStr("data"));
        if (map == null) {
            return;
        }

        // ArrayList<Map<String, Object>> rows = CollUtil.newArrayList(map);

        // 写入Excel
        ExcelWriter writer = ExcelUtil.getWriter(true);
        Workbook workbook = writer.getWorkbook();
        Sheet sheet = workbook.getSheetAt(0);
        CellRangeAddress cra;
        Row row;
        Cell cell;
        CellStyle styleBold;
        CellStyle style;
        Font font;

        // ----------------------创建第一行---------------
        // 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        row = sheet.createRow(0);
        // 创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
        cell = row.createCell(0);
        // 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        cra = new CellRangeAddress(0, 0, 0, 4);
        sheet.addMergedRegion(cra);
        setRegionBorder(BorderStyle.THIN, cra, sheet);

        // 设置单元格内容
        cell.setCellValue("酒店经营信息");

        // 设置样式
        sheet.setDefaultColumnWidth((short) 32);
        font = workbook.createFont();
        styleBold = workbook.createCellStyle();
        style = workbook.createCellStyle();
        font.setFontName("宋体");
        font.setFontHeight((short) 230);
        font.setBold(true);
        styleBold.setFont(font);
        styleBold.setAlignment(HorizontalAlignment.CENTER);
        styleBold.setVerticalAlignment(VerticalAlignment.CENTER);
        styleBold.setBorderBottom(BorderStyle.THIN);
        styleBold.setBorderLeft(BorderStyle.THIN);
        styleBold.setBorderTop(BorderStyle.THIN);
        styleBold.setBorderRight(BorderStyle.THIN);

        // 全局边框
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        cell.setCellStyle(styleBold);

        /**
         *
         * 基本信息
         *
         */
        // ----------------------创建第二行---------------
        row = sheet.createRow(1);
        cra = new CellRangeAddress(1, 11, 0, 0);
        sheet.addMergedRegion(cra);
        setRegionBorder(BorderStyle.THIN, cra, sheet);

        cell = row.createCell(0);
        cell.setCellStyle(styleBold);
        cell.setCellValue("基本信息");

        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("进件申请号码：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("incomeNo") == null ? "" : map.get("incomeNo").toString());

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("统一社会信用代码：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("companyCreditCode") == null ? "" : map.get("companyCreditCode").toString());

        // ----------------------创建第三行---------------
        row = sheet.createRow(2);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("企业名称：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("companyName") == null ? "" : map.get("companyName").toString());

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("酒店名称：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("hotelName") == null ? "" : map.get("hotelName").toString());

        // ----------------------创建第四行---------------
        row = sheet.createRow(3);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("实际控制人名称：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("actualConName") == null ? "" : map.get("actualConName").toString());

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("酒店类型：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        String type = map.get("type") == null ? "" : map.get("type").toString();
        if ("1".equals(type)) {
            cell.setCellValue("贷款酒店");
        } else if ("2".equals(type)) {
            cell.setCellValue("关联酒店");
        } else {
            cell.setCellValue("装修酒店");
        }

        // ----------------------创建第五行---------------
        row = sheet.createRow(4);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("联系人：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("contact") == null ? "" : map.get("contact").toString());

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("联系电话：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("contactTel") == null ? "" : map.get("contactTel").toString());

        // ----------------------创建第六行---------------
        row = sheet.createRow(5);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("酒店经营地址：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("hotelAddress") == null ? "" : map.get("hotelAddress").toString());

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("经营场所性质：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        String campNature = map.get("campNature") == null ? "" : map.get("campNature").toString();
        if ("1".equals(campNature)) {
            cell.setCellValue("租赁");
        } else {
            cell.setCellValue("自有");
        }

        // ----------------------创建第七行---------------
        row = sheet.createRow(6);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("成立年限：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("estYear") == null ? "" : map.get("estYear").toString());

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("是否还清：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        String isPayOff = map.get("isPayOff") == null ? "" : map.get("isPayOff").toString();
        if ("1".equals(isPayOff)) {
            cell.setCellValue("已还清");
        } else {
            cell.setCellValue("未还清");
        }

        // ----------------------创建第八行---------------
        row = sheet.createRow(7);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("消防验收合格证：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        String fireLicenseUrl = map.get("fireLicenseUrl") == null ? "" : map.get("fireLicenseUrl").toString();
        if (!"".equals(fireLicenseUrl)) {
            cell.setCellValue("有");
        } else {
            cell.setCellValue("无");
        }

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("特种行业许可证：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        String specialLicenseUrl = map.get("specialLicenseUrl") == null ? "" : map.get("specialLicenseUrl").toString();
        if (!"".equals(specialLicenseUrl)) {
            cell.setCellValue("有");
        } else {
            cell.setCellValue("无");
        }

        // ----------------------创建第九行---------------
        row = sheet.createRow(8);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("卫生许可证：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        String healthLicenseUrl = map.get("healthLicenseUrl") == null ? "" : map.get("healthLicenseUrl").toString();
        if (!"".equals(healthLicenseUrl)) {
            cell.setCellValue("有");
        } else {
            cell.setCellValue("无");
        }

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("营业执照：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        String businessLicenseUrl = map.get("businessLicenseUrl") == null ? "" : map.get("businessLicenseUrl").toString();
        if (!"".equals(businessLicenseUrl)) {
            cell.setCellValue("有");
        } else {
            cell.setCellValue("无");
        }

        // ----------------------创建第十行---------------
        row = sheet.createRow(9);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("近两年年均营业收入：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("operateIncome") == null ? "" : map.get("operateIncome").toString());

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("在携程近两年年均订房收入：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("ctripIncome") == null ? "" : map.get("ctripIncome").toString());

        // ----------------------创建第十一行---------------
        row = sheet.createRow(10);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("近一年入住率：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("occupancyRate") == null ? "" : map.get("occupancyRate").toString());

        cell = row.createCell(3);
        cell.setCellStyle(style);
        cell.setCellValue("众荟数据酒店经营城市排名：");

        cell = row.createCell(4);
        cell.setCellStyle(style);
        cell.setCellValue(map.get("cityRank") == null ? "" : map.get("cityRank").toString());

        // ----------------------创建第十二行---------------
        row = sheet.createRow(11);
        cell = row.createCell(1);
        cell.setCellStyle(style);
        cell.setCellValue("是否为众荟推荐白名单客户：");

        cell = row.createCell(2);
        cell.setCellStyle(style);
        String isJwisWhitelist = map.get("isJwisWhitelist") == null ? "" : map.get("isJwisWhitelist").toString();
        if ("1".equals(isJwisWhitelist)) {
            cell.setCellValue("是");
        } else {
            cell.setCellValue("否");
        }

        cell = row.createCell(3);
        cell.setCellStyle(style);

        cell = row.createCell(4);
        cell.setCellStyle(style);

        /**
         *
         * 经营数据跟踪
         *
         */
        // 先判断经营数据条数
        List<Map> monthOperateList = (List<Map>) map.get("monthOperateList");
        int opSize;
        if (CollUtil.isEmpty(monthOperateList)) {
            opSize = 0;
        } else {
            opSize = monthOperateList.size();
        }

        // ----------------------创建第十三行---------------
        if (opSize > 0) {
            row = sheet.createRow(12);
            cra = new CellRangeAddress(12, (12 + (opSize * 6 - 1)) - 1, 0, 0);
            sheet.addMergedRegion(cra);
            setRegionBorder(BorderStyle.THIN, cra, sheet);

            cell = row.createCell(0);
            cell.setCellValue("经营数据跟踪");
            cell.setCellStyle(styleBold);

            for (int i = 0; i < opSize; i++) {
                Map operateMap = monthOperateList.get(i);
                if (operateMap != null) {
                    if (i > 0) {
                        row = sheet.createRow(12 + (i * 6));
                    }
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("月份：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(operateMap.get("dataMonth") == null ? "" : operateMap.get("dataMonth").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);

                    cell = row.createCell(4);
                    cell.setCellStyle(style);

                    row = sheet.createRow(12 + (i * 6) + 1);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("当月营业收入：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(operateMap.get("monthIncome") == null ? "" : operateMap.get("monthIncome").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("当月在携程订房收入：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    cell.setCellValue(operateMap.get("monthIncomeFromCtrip") == null ? "" : operateMap.get("monthIncomeFromCtrip").toString());

                    row = sheet.createRow(12 + (i * 6) + 2);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("当月入住率：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(operateMap.get("monthOccupancyRate") == null ? "" : operateMap.get("monthOccupancyRate").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("当月城市排名：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    cell.setCellValue(operateMap.get("monthCityRank") == null ? "" : operateMap.get("monthCityRank").toString());

                    row = sheet.createRow(12 + (i * 6) + 3);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("近12个月营业收入汇总值：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(operateMap.get("operaIncomeSum12") == null ? "" : operateMap.get("operaIncomeSum12").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("近13-24个月营业收入汇总值：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    cell.setCellValue(operateMap.get("operaIncomeSum24") == null ? "" : operateMap.get("operaIncomeSum24").toString());

                    row = sheet.createRow(12 + (i * 6) + 4);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("T-2月酒店入住率：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(operateMap.get("hotelOccupancyRate") == null ? "" : operateMap.get("hotelOccupancyRate").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);

                    cell = row.createCell(4);
                    cell.setCellStyle(style);

                    if (opSize > 1 && i != opSize - 1) {
                        row = sheet.createRow(12 + (i * 6) + 5);
                        cra = new CellRangeAddress(12 + (i * 6) + 5, 12 + (i * 6) + 5, 1, 4);
                        sheet.addMergedRegion(cra);
                        setRegionBorder(BorderStyle.THIN, cra, sheet);
                    }
                }
            }
        }


        /**
         *
         * 加盟信息
         *
         */
        // 先判断加盟数据条数
        List<Map> franchiseList = (List<Map>) map.get("franchiseList");
        int frSize;
        int frBase;
        if (CollUtil.isEmpty(franchiseList)) {
            frSize = 0;
        } else {
            frSize = franchiseList.size();
        }

        if (opSize > 0) {
            frBase = 12 + (opSize * 6 - 1);
        } else {
            frBase = 12;
        }

        if (frSize > 0) {
            row = sheet.createRow(frBase);
            cra = new CellRangeAddress(frBase, (frBase + (frSize * 4 - 1)) - 1, 0, 0);
            sheet.addMergedRegion(cra);
            setRegionBorder(BorderStyle.THIN, cra, sheet);

            cell = row.createCell(0);
            cell.setCellValue("加盟信息");
            cell.setCellStyle(styleBold);

            for (int i = 0; i < frSize; i++) {
                Map franchMap = franchiseList.get(i);
                if (franchMap != null) {
                    if (i > 0) {
                        row = sheet.createRow(frBase + (i * 4));
                    }
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("更新日期：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(franchMap.get("updateTime") == null ? "" : franchMap.get("updateTime").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);

                    cell = row.createCell(4);
                    cell.setCellStyle(style);

                    row = sheet.createRow(frBase + (i * 4) + 1);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("加盟酒店品牌：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(franchMap.get("fchBrand") == null ? "" : franchMap.get("fchBrand").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("加盟日期：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    cell.setCellValue(franchMap.get("fchDate") == null ? "" : franchMap.get("fchDate").toString());

                    row = sheet.createRow(frBase + (i * 4) + 2);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("加盟合同期限：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(franchMap.get("fchContrDeadline") == null ? "" : franchMap.get("fchContrDeadline").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("加盟店面积（㎡）：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    cell.setCellValue(franchMap.get("fchArea") == null ? "" : franchMap.get("fchArea").toString());

                    if (frSize > 1 && i != frSize - 1) {
                        row = sheet.createRow(frBase + (i * 4) + 3);
                        cra = new CellRangeAddress(frBase + (i * 4) + 3, frBase + (i * 4) + 3, 1, 4);
                        sheet.addMergedRegion(cra);
                        setRegionBorder(BorderStyle.THIN, cra, sheet);
                    }
                }
            }
        }


        /**
         *
         * 租赁信息
         *
         */
        // 先判断租赁信息条数
        List<Map> leaseList = (List<Map>) map.get("leaseList");
        int leSize;
        int leBase;
        if (CollUtil.isEmpty(leaseList)) {
            leSize = 0;
        } else {
            leSize = leaseList.size();
        }

        if (frSize > 0) {
            leBase = frBase + (frSize * 4 - 1);
        } else {
            leBase = frBase;
        }

        if (leSize > 0) {
            row = sheet.createRow(leBase);
            cra = new CellRangeAddress(leBase, (leBase + (leSize * 6 - 1)) - 1, 0, 0);
            sheet.addMergedRegion(cra);
            setRegionBorder(BorderStyle.THIN, cra, sheet);

            cell = row.createCell(0);
            cell.setCellValue("租赁合同信息");
            cell.setCellStyle(styleBold);

            for (int i = 0; i < leSize; i++) {
                Map leaseMap = leaseList.get(i);
                if (leaseMap != null) {
                    if (i > 0) {
                        row = sheet.createRow(leBase + (i * 6));
                    }
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("更新日期：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(leaseMap.get("updateTime") == null ? "" : leaseMap.get("updateTime").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);

                    cell = row.createCell(4);
                    cell.setCellStyle(style);

                    row = sheet.createRow(leBase + (i * 6) + 1);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("出租人：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(leaseMap.get("lessor") == null ? "" : leaseMap.get("lessor").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("承租人：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    cell.setCellValue(leaseMap.get("lessee") == null ? "" : leaseMap.get("lessee").toString());

                    row = sheet.createRow(leBase + (i * 6) + 2);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("租赁开始时间：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(leaseMap.get("leaseBeginDate") == null ? "" : leaseMap.get("leaseBeginDate").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("租赁结束时间：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    cell.setCellValue(leaseMap.get("leaseEndDate") == null ? "" : leaseMap.get("leaseEndDate").toString());

                    row = sheet.createRow(leBase + (i * 6) + 3);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("租金：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(leaseMap.get("rent") == null ? "" : leaseMap.get("rent").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("是否可商用：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    String isCommercial = leaseMap.get("isCommercial") == null ? "" : leaseMap.get("isCommercial").toString();
                    if ("1".equals(isCommercial)) {
                        cell.setCellValue("是");
                    } else {
                        cell.setCellValue("否");
                    }

                    row = sheet.createRow(leBase + (i * 6) + 4);
                    cell = row.createCell(1);
                    cell.setCellStyle(style);
                    cell.setCellValue("租赁面积（㎡）：");

                    cell = row.createCell(2);
                    cell.setCellStyle(style);
                    cell.setCellValue(leaseMap.get("leaseArea") == null ? "" : leaseMap.get("leaseArea").toString());

                    cell = row.createCell(3);
                    cell.setCellStyle(style);
                    cell.setCellValue("租赁合同：");

                    cell = row.createCell(4);
                    cell.setCellStyle(style);
                    String leaseContrImgUrl = leaseMap.get("leaseContrImgUrl") == null ? "" : leaseMap.get("leaseContrImgUrl").toString();
                    if (!"".equals(leaseContrImgUrl)) {
                        cell.setCellValue("有");
                    } else {
                        cell.setCellValue("无");
                    }

                    if (leSize > 1 && i != leSize - 1) {
                        row = sheet.createRow(leBase + (i * 6) + 5);
                        cra = new CellRangeAddress(leBase + (i * 6) + 5, leBase + (i * 6) + 5, 1, 4);
                        sheet.addMergedRegion(cra);
                        setRegionBorder(BorderStyle.THIN, cra, sheet);
                    }
                }
            }
        }

        // 制表信息
        int tableBase;
        if (leBase > 0) {
            tableBase = leBase + (leSize * 6 - 1);;
        } else {
            tableBase = leBase;
        }

        row = sheet.createRow(tableBase);
        cra = new CellRangeAddress(tableBase, tableBase, 0, 4);
        sheet.addMergedRegion(cra);
        setRegionBorder(BorderStyle.THIN, cra, sheet);

        cell = row.createCell(0);
        String []today = DateUtil.today().split("-");
        cell.setCellValue("制表时间：" + today[0] + "年" + today[1] + "月" + today[2] + "日              制表人：" + ShiroUtils.getUserEntity().getUsername());
        cell.setCellStyle(styleBold);

        // writer.write(rows);

        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");

        String fileName = URLEncoder.encode("酒店经营数据" + DateUtil.date(), "UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");

        ServletOutputStream out = response.getOutputStream();
        writer.flush(out);
        writer.close();
        IoUtil.close(out);

    }

    /**
     * 设置边框
     *
     * @param border
     * @param region
     * @param sheet
     */
    private void setRegionBorder(BorderStyle border, CellRangeAddress region, Sheet sheet) {
        RegionUtil.setBorderBottom(border, region, sheet);
        RegionUtil.setBorderLeft(border, region, sheet);
        RegionUtil.setBorderRight(border, region, sheet);
        RegionUtil.setBorderTop(border, region, sheet);
    }

    /**
     * 账户余额预警列表
     *
     * @param incomeNo      进件编号
     * @param loanNo        贷款编号
     * @param companyName   企业名称
     * @param accountNumber 账户号
     * @param warnLevel     预警等级
     * @param startDate     应还开始日期
     * @param endDate       应还结束日期
     * @return
     */
    @RequestMapping("/accountBalanceWarn")
    @RequiresPermissions("api:loanAfterManage:accountBalanceWarn")
    public String accountBalanceWarn(@RequestParam(required = false) String incomeNo,
                                     @RequestParam(required = false) String loanNo,
                                     @RequestParam(required = false) String companyName,
                                     @RequestParam(required = false) String accountNumber,
                                     @RequestParam(required = false) Integer warnLevel,
                                     @RequestParam(required = false) String startDate,
                                     @RequestParam(required = false) String endDate,
                                     @RequestParam Integer page,
                                     @RequestParam Integer limit) {
        Map<String, Object> params = new HashMap<>(9);
        params.put("incomeNo", incomeNo);
        params.put("loanNo", loanNo);
        params.put("companyName", companyName);
        params.put("accountNumber", accountNumber);
        params.put("warnLevel", warnLevel);
        params.put("startDate", startDate);
        params.put("endDate", endDate);
        params.put("page", page.toString());
        params.put("limit", limit.toString());
        return HttpUtil.get(apiHost + accountBalanceWarn, params);
    }

    /**
     * 手动更新账户余额跟预警信息
     *
     * @return
     */
    @RequestMapping("/queryAccountBalance")
    @RequiresPermissions("api:loanAfterManage:queryAccountBalance")
    public String queryAccountBalance() {
        String noticeDaysWarn = configService.getValue(SmsConstant.NOTICE_DAYS_WARN);
        Map<String, Object> params = new HashMap<>(1);
        params.put("noticeDaysWarn", noticeDaysWarn);
        return HttpUtil.get(apiHost + queryAccountBalance, params);
    }

    /**
     * 查询还清接口
     *
     * @return
     */
    @RequestMapping("/queryPayOff")
    @RequiresPermissions("api:loanAfterManage:queryPayOff")
    public String queryPayOff(@RequestParam(required = false) String incomeNo) {
        Map<String, Object> params = new HashMap<>(1);
        params.put("incomeNo", incomeNo);
        return HttpUtil.get(apiHost + queryPayOff, params);
    }

}
