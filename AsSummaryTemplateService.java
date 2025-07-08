/**
 * Copyright (c) 2013-Now http://jeesite.com All rights reserved.
 */
package com.jeesite.modules.assummary.service;


import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.date.DateUtil;
import com.jeesite.common.idgen.IdGen;
import com.jeesite.common.service.ServiceException;
import com.jeesite.common.utils.Globals;
import com.jeesite.common.lang.StringUtils;
import com.jeesite.modules.asconfig.enums.ConfigPageCode;
import com.jeesite.modules.assummary.converter.TemplateConverter;
import com.jeesite.modules.assummary.dao.AsSummaryTemplateFieldDao;
import com.jeesite.modules.assummary.dao.AsSummaryTemplateOptionDao;
import com.jeesite.modules.assummary.dao.AsSummaryTemplateTableColumnDao;
import com.jeesite.modules.assummary.entity.AsSummaryTemplateField;
import com.jeesite.modules.assummary.entity.AsSummaryTemplateOption;
import com.jeesite.modules.assummary.entity.AsSummaryTemplateTableColumn;
import com.jeesite.modules.assummary.vo.req.AsSummaryTemplateActionDTO;
import com.jeesite.modules.assummary.vo.req.AsSummaryTemplateSave;
import com.jeesite.modules.assummary.vo.req.frontend.TemplateFormDTO;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;


import com.jeesite.common.entity.Page;
import com.jeesite.common.service.CrudService;
import com.jeesite.modules.assummary.entity.AsSummaryTemplate;
import com.jeesite.modules.assummary.dao.AsSummaryTemplateDao;
import com.jeesite.modules.assummary.vo.req.AsSummaryTemplateQueryDTO;
import com.jeesite.modules.assummary.vo.res.AsSummaryTemplatePageVO;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;

import java.util.*;
import java.util.stream.Collectors;

import com.jeesite.modules.asconfig.dao.AsColumnConfigDetailDao;
import com.jeesite.modules.asconfig.entity.AsColumnConfigDetail;
import com.jeesite.modules.assummary.vo.res.AsSummaryTemplateColumnVO;
import com.jeesite.modules.assummary.vo.res.AsSummaryTemplateDetailVO;
import com.jeesite.modules.assummary.vo.req.AsSummaryTemplateDetailQueryDTO;
import com.jeesite.modules.assummary.vo.res.FieldGroupVO;
import com.jeesite.modules.assummary.vo.res.FieldContentVO;
import com.jeesite.modules.assummary.vo.res.FieldReplyVO;
import com.jeesite.modules.assummary.vo.res.FieldDataVO;
import com.jeesite.modules.assummary.vo.res.OptionItemVO;
import com.jeesite.modules.assummary.vo.res.TextOptionVO;
import com.jeesite.modules.assummary.vo.res.TableColumnVO;
import com.jeesite.modules.assummary.vo.excel.TemplateFieldNode;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

/**
 * 年度总结模板主表Service
 *
 * @author wt
 * @version 2025-07-03
 */
@Service
@Transactional(readOnly = true)
public class AsSummaryTemplateService extends CrudService<AsSummaryTemplateDao, AsSummaryTemplate> {

    @Autowired
    private AsSummaryTemplateDao asSummaryTemplateDao;

    @Autowired
    private AsSummaryTemplateFieldDao asSummaryTemplateFieldDao;

    @Autowired
    private AsSummaryTemplateOptionDao asSummaryTemplateOptionDao;

    @Autowired
    private AsSummaryTemplateTableColumnDao asSummaryTemplateTableColumnDao;

    @Autowired
    private AsColumnConfigDetailDao asColumnConfigDetailDao;

    // 添加表格起始列常量和当前填表说明列索引
    private static final int TABLE_START_COL = 2;
    private int fillReasonCol;

    /**
     * 获取单条数据
     *
     * @param asSummaryTemplate
     * @return
     */
    @Override
    public AsSummaryTemplate get(AsSummaryTemplate asSummaryTemplate) {
        return super.get(asSummaryTemplate);
    }

    /**
     * 查询分页数据
     *
     * @param page              分页对象
     * @param asSummaryTemplate
     * @return
     */
    @Override
    public Page<AsSummaryTemplate> findPage(Page<AsSummaryTemplate> page, AsSummaryTemplate asSummaryTemplate) {
        return super.findPage(page, asSummaryTemplate);
    }

    /**
     * 保存数据（仅插入，不做更新）
     *
     * @param asSummaryTemplateSave
     * @return 新创建的模板ID
     */
    @Transactional(rollbackFor = Exception.class)
    public String save(AsSummaryTemplateSave asSummaryTemplateSave) {
        String tablePrefix = Globals.getTableFix(asSummaryTemplateSave.getUnitCode(), asSummaryTemplateSave.getHospitalCode());

        // 生成新的模板ID并保存主表
        String id = IdGen.nextId();
        asSummaryTemplateSave.setId(id);

        AsSummaryTemplate tpl = new AsSummaryTemplate();
        BeanUtil.copyProperties(asSummaryTemplateSave, tpl);
        tpl.setTablePrifix(tablePrefix);
        tpl.setId(id);
        tpl.setInitType(1L);
        tpl.setCreateUser(asSummaryTemplateSave.getUserId());
        tpl.setCreateTime(DateUtil.date());
        tpl.setUpdateUser(asSummaryTemplateSave.getUserId());
        tpl.setUpdateTime(DateUtil.date());
        asSummaryTemplateDao.insert(tpl);

        logger.info("已创建新模板[{}]: {}", id, tpl.getName());

        // 2. 批量保存字段数据
        saveFields(asSummaryTemplateSave, tablePrefix, id);

        return id;
    }

    /**
     * 更新数据
     *
     * @param asSummaryTemplateSave
     * @return 更新的模板ID
     */
    @Transactional(rollbackFor = Exception.class)
    public String update(AsSummaryTemplateSave asSummaryTemplateSave) {
        if (StringUtils.isBlank(asSummaryTemplateSave.getId())) {
            throw new ServiceException("模板ID不能为空");
        }

        String tablePrefix = Globals.getTableFix(asSummaryTemplateSave.getUnitCode(), asSummaryTemplateSave.getHospitalCode());

        // 检查模板是否存在
        AsSummaryTemplate template = asSummaryTemplateDao.getTemplateById(tablePrefix, asSummaryTemplateSave.getId());
        if (template == null) {
            throw new ServiceException("模板不存在，无法更新");
        }
        if (Objects.equals(template.getInitType(), Globals.COMMON_LONG_0)) {
            throw new ServiceException("初始化模板不允许删除");
        }

        // 更新主表数据
        AsSummaryTemplate tpl = new AsSummaryTemplate();
        BeanUtil.copyProperties(asSummaryTemplateSave, tpl);
        tpl.setTablePrifix(tablePrefix);
        tpl.setUpdateUser(asSummaryTemplateSave.getUserId());
        tpl.setUpdateTime(DateUtil.date());
        asSummaryTemplateDao.update(tpl);

        logger.info("已更新模板[{}]主表数据", asSummaryTemplateSave.getId());

        // 清理所有子表数据
        deleteRelatedData(tablePrefix, asSummaryTemplateSave.getId(), asSummaryTemplateSave.getUserId());

        // 批量保存字段数据
        saveFields(asSummaryTemplateSave, tablePrefix, asSummaryTemplateSave.getId());

        return asSummaryTemplateSave.getId();
    }

    /**
     * 保存字段及相关数据
     *
     * @param templateSave 模板保存对象
     * @param tablePrefix  表前缀
     * @param templateId   模板ID
     */
    private void saveFields(AsSummaryTemplateSave templateSave, String tablePrefix, String templateId) {
        try {
            // 预先计算字段数量用于日志记录
            final int fieldSize = templateSave.getFields().size();

            // 批量保存字段数据
            List<AsSummaryTemplateField> fieldList = templateSave.getFields().stream()
                    .map(f -> {
                        // 构造字段实体对象
                        AsSummaryTemplateField field = new AsSummaryTemplateField();
                        BeanUtil.copyProperties(f, field);
                        field.setTablePrifix(tablePrefix);
                        field.setId(f.getId());
                        field.setTemplateId(templateId);
                        field.setIsRequired(Objects.isNull(f.getRequired()) ? 0 : f.getRequired() ? 1 : 0);
                        field.setAllowUploadImage(Objects.isNull(f.getAllowImage()) ? 0 : f.getAllowImage() ? 1 : 0);
                        field.setCreateUser(templateSave.getUserId());
                        field.setCreateTime(DateUtil.date());
                        return field;
                    })
                    .collect(Collectors.toList());

            // 批量保存字段数据
            if (!fieldList.isEmpty()) {
                int insertCount = asSummaryTemplateFieldDao.batchInsert(tablePrefix, fieldList);
                logger.info("批量保存模板[{}]字段数据: {}条，插入成功: {}条", templateId, fieldSize, insertCount);
            }

            // 批量保存选项数据
            List<AsSummaryTemplateOption> optionList = templateSave.getFields().stream()
                    .filter(f -> f.getOptions() != null && !f.getOptions().isEmpty())
                    .flatMap(f -> f.getOptions().stream().map(o -> {
                        AsSummaryTemplateOption opt = new AsSummaryTemplateOption();
                        BeanUtil.copyProperties(o, opt);
                        opt.setTablePrifix(tablePrefix);
                        opt.setId(IdGen.nextId());
                        opt.setFieldId(f.getId());
                        opt.setCreateUser(templateSave.getUserId());
                        opt.setCreateTime(DateUtil.date());
                        return opt;
                    }))
                    .collect(Collectors.toList());

            // 批量保存选项数据
            if (!optionList.isEmpty()) {
                int insertCount = asSummaryTemplateOptionDao.batchInsert(tablePrefix, optionList);
                logger.info("批量保存模板[{}]选项数据: {}条，插入成功: {}条", templateId, optionList.size(), insertCount);
            }

            // 批量保存表格列数据
            List<AsSummaryTemplateTableColumn> columnList = templateSave.getFields().stream()
                    .filter(f -> "TABLE".equalsIgnoreCase(f.getFieldType()))
                    .filter(f -> f.getColumns() != null && !f.getColumns().isEmpty())
                    .flatMap(f -> f.getColumns().stream().map(c -> {
                        AsSummaryTemplateTableColumn col = new AsSummaryTemplateTableColumn();
                        BeanUtil.copyProperties(c, col);
                        col.setTablePrifix(tablePrefix);
                        col.setId(IdGen.nextId());
                        col.setFieldId(f.getId());
                        col.setCreateUser(templateSave.getUserId());
                        col.setCreateTime(DateUtil.date());
                        return col;
                    }))
                    .collect(Collectors.toList());

            // 批量保存表格列数据
            if (!columnList.isEmpty()) {
                int insertCount = asSummaryTemplateTableColumnDao.batchInsert(tablePrefix, columnList);
                logger.info("批量保存模板[{}]表格列数据: {}条，插入成功: {}条", templateId, columnList.size(), insertCount);
            }
        } catch (Exception e) {
            logger.error("保存模板字段数据时发生错误: {}", e.getMessage(), e);
            throw new ServiceException("保存模板字段数据时发生错误: " + e.getMessage(), e);
        }
    }

    /**
     * 删除模板相关的子表数据（使用逻辑删除）
     *
     * @param tablePrefix 表前缀
     * @param templateId  模板ID
     * @param updateUser  更新人
     */
    private void deleteRelatedData(String tablePrefix, String templateId, String updateUser) {
        try {
            // 获取模板关联的所有字段ID
            List<String> fieldIds = asSummaryTemplateFieldDao.findFieldIdsByTemplateId(tablePrefix, templateId);

            // 如果没有关联字段，快速返回
            if (fieldIds == null || fieldIds.isEmpty()) {
                logger.info("模板[{}]无关联字段数据，无需执行删除操作", templateId);
                return;
            }

            logger.info("模板[{}]关联的字段数量: {}", templateId, fieldIds.size());

            // 批量逻辑删除关联的选项数据
            int optionUpdateCount = asSummaryTemplateOptionDao.batchLogicDeleteByFieldIds(tablePrefix, fieldIds, updateUser);
            logger.info("批量逻辑删除模板[{}]选项数据: {}条", templateId, optionUpdateCount);

            // 批量逻辑删除关联的表格列数据
            int columnUpdateCount = asSummaryTemplateTableColumnDao.batchLogicDeleteByFieldIds(tablePrefix, fieldIds, updateUser);
            logger.info("批量逻辑删除模板[{}]表格列数据: {}条", templateId, columnUpdateCount);

            // 批量逻辑删除字段数据
            int fieldUpdateCount = asSummaryTemplateFieldDao.logicDeleteByTemplateId(tablePrefix, templateId, updateUser);
            logger.info("批量逻辑删除模板[{}]字段数据: {}条", templateId, fieldUpdateCount);

        } catch (Exception e) {
            logger.error("逻辑删除模板相关数据时发生错误: {}", e.getMessage(), e);
            throw new ServiceException("逻辑删除模板相关数据时发生错误: " + e.getMessage(), e);
        }
    }

    /**
     * 从前端表单保存模板数据（新建）
     *
     * @param templateForm 前端表单数据
     * @return 保存的模板ID
     */
    @Transactional(rollbackFor = Exception.class)
    public String saveFromFrontend(TemplateFormDTO templateForm) {
        // 使用转换器将前端数据转换为后端数据结构
        AsSummaryTemplateSave templateSave = TemplateConverter.convertToTemplateSave(templateForm);

        // 调用保存方法
        return save(templateSave);
    }

    /**
     * 从前端表单更新模板数据
     *
     * @param templateForm 前端表单数据
     * @return 更新的模板ID
     */
    @Transactional(rollbackFor = Exception.class)
    public String updateFromFrontend(TemplateFormDTO templateForm) {
        // 检查ID是否存在
        if (StringUtils.isBlank(templateForm.getId())) {
            throw new ServiceException("模板ID不能为空，无法更新");
        }

        logger.info("开始更新模板[{}]：{}", templateForm.getId(), templateForm.getName());

        // 使用转换器将前端数据转换为后端数据结构，并确保设置ID
        AsSummaryTemplateSave templateSave = TemplateConverter.convertToTemplateSave(templateForm);
        templateSave.setId(templateForm.getId());

        // 调用更新方法
        String id = update(templateSave);

        logger.info("模板[{}]更新完成", id);
        return id;
    }

    /**
     * 更新状态
     *
     * @param asSummaryTemplate
     */
    @Override
    @Transactional(readOnly = false)
    public void updateStatus(AsSummaryTemplate asSummaryTemplate) {
        super.updateStatus(asSummaryTemplate);
    }

    /**
     * 删除数据
     *
     * @param asSummaryTemplate
     */
    @Override
    @Transactional(readOnly = false)
    public void delete(AsSummaryTemplate asSummaryTemplate) {
        super.delete(asSummaryTemplate);
    }

    /**
     * 查询模板列表
     *
     * @param queryDTO 查询条件
     * @return 分页模板列表
     */
    public AsSummaryTemplatePageVO findTemplateList(AsSummaryTemplateQueryDTO queryDTO) {
        try {
            // 获取表前缀
            String tablePrefix = Globals.getTableFix(queryDTO.getUnitCode(), queryDTO.getHospitalCode());

            // 设置分页参数
            Integer pageNo = queryDTO.getCurrentPage();
            Integer pageSize = queryDTO.getPageSize();

            // 获取列配置
            List<AsColumnConfigDetail> columnConfigs = getColumnConfigs(tablePrefix, ConfigPageCode.TEMPLATE.getCode());

            // 提取需要查询的列名
            List<String> columnKeys = columnConfigs.stream()
                    .map(AsColumnConfigDetail::getColumnKey)
                    .collect(Collectors.toList());

            // 使用PageHelper进行分页
            PageHelper.startPage(pageNo, pageSize);
            List<Map<String, Object>> templateList = asSummaryTemplateDao.findTemplateList(
                    tablePrefix,
                    queryDTO.getName(),
                    columnKeys
            );

            // 获取分页信息
            PageInfo<Map<String, Object>> pageInfo = new PageInfo<>(templateList);

            // 转换列配置为前端所需格式
            List<AsSummaryTemplateColumnVO> columns = columnConfigs.stream().map(config -> {
                AsSummaryTemplateColumnVO columnVO = new AsSummaryTemplateColumnVO();
                columnVO.setColumnKey(config.getColumnKey());
                columnVO.setColumnLabel(config.getColumnLabel());
                columnVO.setColumnOrder(config.getColumnOrder() != null ? config.getColumnOrder().intValue() : 0);
                columnVO.setColumnWidth(config.getColumnWidth() != null ? config.getColumnWidth().intValue() : null);
                columnVO.setVisible(Objects.equals(config.getIsVisible(), 1));
                return columnVO;
            }).collect(Collectors.toList());

            // 构建返回结果
            AsSummaryTemplatePageVO pageVO = new AsSummaryTemplatePageVO();
            pageVO.setPageNo(pageNo);
            pageVO.setPageSize(pageSize);
            pageVO.setTotal(pageInfo.getTotal());
            pageVO.setTotalPages(pageInfo.getPages());
            pageVO.setColumns(columns);
            pageVO.setList(templateList);

            return pageVO;
        } catch (Exception e) {
            logger.error("查询模板列表失败：{}", e.getMessage(), e);
            throw new ServiceException("查询模板列表失败：" + e.getMessage(), e);
        }
    }

    /**
     * 获取列配置
     *
     * @param tablePrefix 表前缀
     * @param pageCode    页面编码
     * @return 列配置列表
     */
    private List<AsColumnConfigDetail> getColumnConfigs(String tablePrefix, String pageCode) {
        // 先查询用户自定义配置
        List<AsColumnConfigDetail> customConfigs = asColumnConfigDetailDao.findCustomColumnConfig(tablePrefix, pageCode, Globals.COMMON_STRING_0);

        // 如果有自定义配置，则返回自定义配置
        if (customConfigs != null && !customConfigs.isEmpty()) {
            return customConfigs;
        }

        // 否则返回默认配置
        return asColumnConfigDetailDao.findCustomColumnConfigByPageCode(tablePrefix, pageCode, Globals.COMMON_STRING_1);
    }

    /**
     * 查询模板详情
     *
     * @param queryDTO 查询条件
     * @return 模板详情
     */
    public AsSummaryTemplateDetailVO findTemplateDetail(AsSummaryTemplateDetailQueryDTO queryDTO) {
        try {
            // 获取表前缀
            String tablePrefix = Globals.getTableFix(queryDTO.getUnitCode(), queryDTO.getHospitalCode());

            // 查询模板基本信息
            AsSummaryTemplate template = asSummaryTemplateDao.getTemplateById(tablePrefix, queryDTO.getId());
            if (template == null) {
                throw new ServiceException("模板不存在");
            }

            // 查询模板所有字段
            List<AsSummaryTemplateField> allFields = asSummaryTemplateFieldDao.findFieldsByTemplateId(tablePrefix, queryDTO.getId());
            if (allFields == null || allFields.isEmpty()) {
                logger.warn("模板[{}]没有配置任何字段", queryDTO.getId());
                // 返回基本信息
                AsSummaryTemplateDetailVO detailVO = new AsSummaryTemplateDetailVO();
                detailVO.setName(template.getName());
                detailVO.setDescription(template.getDescription());
                detailVO.setFields(new ArrayList<>());
                return detailVO;
            }

            // 收集所有字段ID
            List<String> allFieldIds = allFields.stream()
                    .map(AsSummaryTemplateField::getId)
                    .collect(Collectors.toList());

            // 查询所有选项并按字段ID分组
            List<AsSummaryTemplateOption> options = asSummaryTemplateOptionDao.findOptionsByFieldIds(tablePrefix, allFieldIds);
            Map<String, List<AsSummaryTemplateOption>> fieldOptionsMap = options.stream()
                    .collect(Collectors.groupingBy(AsSummaryTemplateOption::getFieldId));

            // 查询所有表格列并按字段ID分组
            List<AsSummaryTemplateTableColumn> columns = asSummaryTemplateTableColumnDao.findColumnsByFieldIds(tablePrefix, allFieldIds);
            Map<String, List<AsSummaryTemplateTableColumn>> fieldColumnsMap = columns.stream()
                    .collect(Collectors.groupingBy(AsSummaryTemplateTableColumn::getFieldId));

            // 构建字段层级关系
            // 找出所有顶级字段（项目，parentFieldId为null的字段）
            List<AsSummaryTemplateField> projectFields = allFields.stream()
                    .filter(field -> StringUtils.isBlank(field.getParentFieldId()))
                    .sorted(Comparator.comparing(AsSummaryTemplateField::getSortOrder))
                    .collect(Collectors.toList());

            // 将所有字段按照parentFieldId进行分组，用于快速查找子字段
            Map<String, List<AsSummaryTemplateField>> fieldsByParentId = allFields.stream()
                    .filter(field -> StringUtils.isNotBlank(field.getParentFieldId()))
                    .collect(Collectors.groupingBy(AsSummaryTemplateField::getParentFieldId));

            // 构建前端需要的格式
            AsSummaryTemplateDetailVO detailVO = new AsSummaryTemplateDetailVO();
            detailVO.setName(template.getName());
            detailVO.setDescription(template.getDescription());

            // 构建字段分组列表（项目级别）
            List<FieldGroupVO> fieldGroups = projectFields.stream()
                    .map(projectField -> {
                        FieldGroupVO groupVO = new FieldGroupVO();
                        groupVO.setName(projectField.getLabel());

                        // 获取该项目下的所有内容字段
                        List<AsSummaryTemplateField> contentFields = fieldsByParentId.getOrDefault(projectField.getId(), new ArrayList<>());
                        contentFields.sort(Comparator.comparing(AsSummaryTemplateField::getSortOrder));

                        // 构建内容列表
                        List<FieldContentVO> contents = contentFields.stream()
                                .map(contentField -> {
                                    // 创建内容VO
                                    FieldContentVO contentVO = new FieldContentVO();
                                    contentVO.setName(contentField.getLabel());
                                    contentVO.setRequired(contentField.getIsRequired() != null && contentField.getIsRequired() == 1);
                                    contentVO.setAllowImage(contentField.getAllowUploadImage() != null && contentField.getAllowUploadImage() == 1);

                                    // 获取该内容下的所有回复字段
                                    List<AsSummaryTemplateField> replyFields = fieldsByParentId.getOrDefault(contentField.getId(), new ArrayList<>());
                                    replyFields.sort(Comparator.comparing(AsSummaryTemplateField::getSortOrder));

                                    // 构建回复列表
                                    List<FieldReplyVO> replies = replyFields.stream()
                                            .map(replyField -> {
                                                FieldReplyVO replyVO = new FieldReplyVO();
                                                replyVO.setFieldType(replyField.getFieldType());

                                                // 构建字段数据
                                                FieldDataVO dataVO = new FieldDataVO();
                                                dataVO.setInstruction(replyField.getDescription());

                                                // 根据字段类型处理不同的数据
                                                switch (replyField.getFieldType()) {
                                                    case "SINGLE_CHOICE":
                                                        // 处理单选项
                                                        List<OptionItemVO> radioOptions = fieldOptionsMap.getOrDefault(replyField.getId(), Collections.emptyList())
                                                                .stream()
                                                                .map(option -> {
                                                                    OptionItemVO itemVO = new OptionItemVO();
                                                                    itemVO.setText(option.getOptionLabel());
                                                                    return itemVO;
                                                                })
                                                                .collect(Collectors.toList());
                                                        dataVO.setRadioOptions(radioOptions);
                                                        break;
                                                    case "TEXT":
                                                        // 处理文本选项
                                                        List<TextOptionVO> textOptions = fieldOptionsMap.getOrDefault(replyField.getId(), Collections.emptyList())
                                                                .stream()
                                                                .map(option -> {
                                                                    TextOptionVO itemVO = new TextOptionVO();
                                                                    itemVO.setText(option.getOptionLabel());
                                                                    itemVO.setPrefixion(replyField.getPrefix() != null ? replyField.getPrefix() : "");
                                                                    itemVO.setPostfix(replyField.getSuffix() != null ? replyField.getSuffix() : "");
                                                                    return itemVO;
                                                                })
                                                                .collect(Collectors.toList());
                                                        dataVO.setTextOptions(textOptions);
                                                        break;
                                                    case "TABLE":
                                                        // 处理表格列
                                                        List<TableColumnVO> tableColumns = fieldColumnsMap.getOrDefault(replyField.getId(), Collections.emptyList())
                                                                .stream()
                                                                .map(col -> {
                                                                    TableColumnVO columnVO = new TableColumnVO();
                                                                    columnVO.setProp(col.getColumnName());
                                                                    columnVO.setLabel(col.getColumnLabel());
                                                                    columnVO.setEditing(false);
                                                                    columnVO.setTempLabel(col.getColumnLabel());
                                                                    return columnVO;
                                                                })
                                                                .collect(Collectors.toList());
                                                        dataVO.setColumns(tableColumns);
                                                        break;
                                                    case "TEXTAREA":
                                                        // 文本域不需要额外数据
                                                        break;
                                                    default:
                                                        logger.warn("未知的字段类型: {}", replyField.getFieldType());
                                                }

                                                replyVO.setData(dataVO);
                                                return replyVO;
                                            })
                                            .collect(Collectors.toList());

                                    contentVO.setReplies(replies);
                                    return contentVO;
                                })
                                .collect(Collectors.toList());

                        groupVO.setContents(contents);
                        return groupVO;
                    })
                    .collect(Collectors.toList());

            detailVO.setFields(fieldGroups);
            return detailVO;

        } catch (Exception e) {
            logger.error("查询模板详情失败：{}", e.getMessage(), e);
            throw new ServiceException("查询模板详情失败：" + e.getMessage(), e);
        }
    }

    /**
     * 复制模板
     *
     * @param actionDTO
     * @return 新模板ID
     */
    @Transactional(rollbackFor = Exception.class)
    public String copyTemplate(AsSummaryTemplateActionDTO actionDTO) {

        String tablePrefix = Globals.getTableFix(actionDTO.getUnitCode(), actionDTO.getHospitalCode());

        // 获取源模板信息
        AsSummaryTemplate sourceTemplate = new AsSummaryTemplate(actionDTO.getTemplateId());
        sourceTemplate.setTablePrifix(tablePrefix);
        sourceTemplate.setDeleteFlag(Globals.COMMON_INTEGER_0);
        sourceTemplate = asSummaryTemplateDao.get(sourceTemplate);
        if (sourceTemplate == null) {
            throw new ServiceException("源模板不存在，无法复制");
        }

        // 创建新模板对象，复制基本信息
        String newTemplateId = IdGen.nextId();
        AsSummaryTemplate newTemplate = new AsSummaryTemplate();
        newTemplate.setTablePrifix(tablePrefix);
        newTemplate.setId(newTemplateId);
        newTemplate.setName(sourceTemplate.getName() + " - 复制");
        newTemplate.setDescription(sourceTemplate.getDescription());
        newTemplate.setInitType(Globals.COMMON_LONG_2);
        newTemplate.setDeleteFlag(Globals.COMMON_INTEGER_0);
        newTemplate.setCreateUser(actionDTO.getUserId());
        newTemplate.setCreateTime(DateUtil.date());
        newTemplate.setUpdateUser(null);
        newTemplate.setUpdateTime(null);

        // 保存新模板基本信息
        asSummaryTemplateDao.insert(newTemplate);
        logger.info("已创建模板副本[{}]: {}", newTemplateId, newTemplate.getName());

        // 查询源模板的所有字段
        List<AsSummaryTemplateField> sourceFields = asSummaryTemplateFieldDao.findFieldsByTemplateId(tablePrefix, actionDTO.getTemplateId());
        if (sourceFields == null || sourceFields.isEmpty()) {
            logger.info("源模板[{}]没有字段，复制完成", actionDTO.getTemplateId());
            return newTemplateId;
        }

        // 创建字段ID映射表，用于维护新旧字段ID的对应关系（用于设置parentFieldId）
        Map<String, String> fieldIdMapping = new HashMap<>();

        // 复制字段
        List<AsSummaryTemplateField> newFields = sourceFields.stream().map(sourceField -> {
            // 生成新字段ID
            String newFieldId = IdGen.nextId();
            // 记录字段ID映射关系
            fieldIdMapping.put(sourceField.getId(), newFieldId);

            // 创建新字段
            AsSummaryTemplateField newField = new AsSummaryTemplateField();
            BeanUtils.copyProperties(sourceField, newField);
            newField.setTablePrifix(tablePrefix);
            newField.setId(newFieldId);
            newField.setTemplateId(newTemplateId);
            newField.setCreateUser(actionDTO.getUserId());
            newField.setCreateTime(DateUtil.date());
            newField.setUpdateUser(null);
            newField.setUpdateTime(null);

            return newField;
        }).collect(Collectors.toList());

        // 更新父字段ID引用
        newFields.forEach(field -> {
            if (StringUtils.isNotBlank(field.getParentFieldId())) {
                field.setParentFieldId(fieldIdMapping.get(field.getParentFieldId()));
            }
        });

        // 批量插入新字段
        if (!newFields.isEmpty()) {
            asSummaryTemplateFieldDao.batchInsert(tablePrefix, newFields);
            logger.info("已复制模板[{}]字段数据: {}条", actionDTO.getTemplateId(), newFields.size());
        }

        // 收集所有源字段ID
        List<String> sourceFieldIds = sourceFields.stream()
                .map(AsSummaryTemplateField::getId)
                .collect(Collectors.toList());

        // 复制选项数据
        List<AsSummaryTemplateOption> sourceOptions = asSummaryTemplateOptionDao.findOptionsByFieldIds(tablePrefix, sourceFieldIds);
        if (!sourceOptions.isEmpty()) {
            List<AsSummaryTemplateOption> newOptions = sourceOptions.stream().map(sourceOption -> {
                // 创建新选项
                AsSummaryTemplateOption newOption = new AsSummaryTemplateOption();
                BeanUtils.copyProperties(sourceOption, newOption);
                newOption.setTablePrifix(tablePrefix);
                newOption.setId(IdGen.nextId());
                newOption.setFieldId(fieldIdMapping.get(sourceOption.getFieldId()));
                newOption.setCreateUser(actionDTO.getUserId());
                newOption.setCreateTime(DateUtil.date());
                newOption.setUpdateUser(null);
                newOption.setUpdateTime(null);

                return newOption;
            }).collect(Collectors.toList());

            // 批量插入新选项
            asSummaryTemplateOptionDao.batchInsert(tablePrefix, newOptions);
            logger.info("已复制模板[{}]选项数据: {}条", actionDTO.getTemplateId(), newOptions.size());
        }

        // 复制表格列数据
        List<AsSummaryTemplateTableColumn> sourceColumns = asSummaryTemplateTableColumnDao.findColumnsByFieldIds(tablePrefix, sourceFieldIds);
        if (!sourceColumns.isEmpty()) {
            List<AsSummaryTemplateTableColumn> newColumns = sourceColumns.stream().map(sourceColumn -> {
                // 创建新表格列
                AsSummaryTemplateTableColumn newColumn = new AsSummaryTemplateTableColumn();
                BeanUtils.copyProperties(sourceColumn, newColumn);
                newColumn.setTablePrifix(tablePrefix);
                newColumn.setId(IdGen.nextId());
                newColumn.setFieldId(fieldIdMapping.get(sourceColumn.getFieldId()));
                newColumn.setCreateUser(actionDTO.getUserId());
                newColumn.setCreateTime(DateUtil.date());
                newColumn.setUpdateUser(null);
                newColumn.setUpdateTime(null);

                return newColumn;
            }).collect(Collectors.toList());

            // 批量插入新表格列
            asSummaryTemplateTableColumnDao.batchInsert(tablePrefix, newColumns);
            logger.info("已复制模板[{}]表格列数据: {}条", actionDTO.getTemplateId(), newColumns.size());
        }

        return newTemplateId;
    }

    /**
     * 删除模板（逻辑删除）
     *
     * @param actionDTO
     * @return 删除结果
     */
    @Transactional(rollbackFor = Exception.class)
    public boolean deleteTemplate(AsSummaryTemplateActionDTO actionDTO) {
        String tablePrefix = Globals.getTableFix(actionDTO.getUnitCode(), actionDTO.getHospitalCode());

        AsSummaryTemplate template = asSummaryTemplateDao.getTemplateById(tablePrefix, actionDTO.getTemplateId());
        if (template == null) {
            throw new ServiceException("模板不存在，无法删除");
        }

        if (Objects.equals(template.getInitType(), Globals.COMMON_LONG_0)) {
            throw new ServiceException("系统初始模板不允许删除");
        }

        Date currentDate = DateUtil.date();
        int notificationCount = asSummaryTemplateDao.countActiveNotificationsByTemplateId(tablePrefix, actionDTO.getTemplateId(), currentDate);
        if (notificationCount > 0) {
            throw new ServiceException("该模板关联的通知中存在未开始或进行中的通知，无法删除");
        }

        try {
            List<String> fieldIds = asSummaryTemplateFieldDao.findFieldIdsByTemplateId(tablePrefix, actionDTO.getTemplateId());

            if (fieldIds != null && !fieldIds.isEmpty()) {
                asSummaryTemplateOptionDao.batchLogicDeleteByFieldIds(tablePrefix, fieldIds, actionDTO.getUserId());

                asSummaryTemplateTableColumnDao.batchLogicDeleteByFieldIds(tablePrefix, fieldIds, actionDTO.getUserId());

                asSummaryTemplateFieldDao.logicDeleteByTemplateId(tablePrefix, actionDTO.getTemplateId(), actionDTO.getUserId());
            }

            template.setDeleteFlag(Globals.COMMON_INTEGER_1);
            template.setUpdateUser(actionDTO.getUserId());
            template.setUpdateTime(DateUtil.date());
            template.setUnitCode(actionDTO.getUnitCode());
            template.setHospitalCode(actionDTO.getHospitalCode());
            asSummaryTemplateDao.update(template);

            logger.info("已成功逻辑删除模板[{}]: {}", actionDTO.getTemplateId(), template.getName());
            return true;
        } catch (Exception e) {
            logger.error("删除模板数据时发生错误: {}", e.getMessage(), e);
            throw new ServiceException("删除模板数据时发生错误: " + e.getMessage(), e);
        }
    }

    /**
     * 导出模板为Excel
     *
     * @param exportDTO 导出请求
     * @return Excel文件字节数组和文件名
     */
    public Map<String, Object> exportTemplateToExcel(AsSummaryTemplateActionDTO exportDTO) {
        try {
            String tablePrefix = Globals.getTableFix(exportDTO.getUnitCode(), exportDTO.getHospitalCode());
            
            AsSummaryTemplate template = asSummaryTemplateDao.getTemplateById(tablePrefix, exportDTO.getTemplateId());
            if (template == null) {
                throw new ServiceException("模板不存在，无法下载");
            }
            
            List<AsSummaryTemplateField> allFields = asSummaryTemplateFieldDao.findFieldsByTemplateId(tablePrefix, exportDTO.getTemplateId());
            if (allFields == null || allFields.isEmpty()) {
                throw new ServiceException("模板没有任何字段定义，无法下载");
            }
            
            // 收集所有字段ID
            List<String> allFieldIds = allFields.stream()
                    .map(AsSummaryTemplateField::getId)
                    .collect(Collectors.toList());
            
            // 查询所有选项并按字段ID分组
            Map<String, List<AsSummaryTemplateOption>> fieldOptionsMap = asSummaryTemplateOptionDao.findOptionsByFieldIds(tablePrefix, allFieldIds)
                    .stream()
                    .collect(Collectors.groupingBy(AsSummaryTemplateOption::getFieldId));
            
            // 查询所有表格列并按字段ID分组
            Map<String, List<AsSummaryTemplateTableColumn>> fieldColumnsMap = asSummaryTemplateTableColumnDao.findColumnsByFieldIds(tablePrefix, allFieldIds)
                    .stream()
                    .collect(Collectors.groupingBy(AsSummaryTemplateTableColumn::getFieldId));
            
            // 构建优化后的字段树，只包含有效显示的节点
            List<TemplateFieldNode> fieldTree = buildOptimizedFieldTree(allFields, fieldOptionsMap, fieldColumnsMap);
            
            byte[] excelBytes = generateExcel(template.getName(), fieldTree);
            
            String timestamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
            String fileName = template.getName() + "_" + timestamp + ".xlsx";
            
            Map<String, Object> result = new HashMap<>();
            result.put("fileName", fileName);
            result.put("fileContent", excelBytes);
            
            return result;
        } catch (Exception e) {
            logger.error("导出模板Excel失败：{}", e.getMessage(), e);
            throw new ServiceException("导出模板Excel失败：" + e.getMessage(), e);
        }
    }
    
    /**
     * 构建优化的字段树状结构，只包含需要显示的节点
     *
     * @param allFields 所有字段
     * @param fieldOptionsMap 字段选项映射
     * @param fieldColumnsMap 字段表格列映射
     * @return 优化后的字段树状结构
     */
    private List<TemplateFieldNode> buildOptimizedFieldTree(
            List<AsSummaryTemplateField> allFields, 
            Map<String, List<AsSummaryTemplateOption>> fieldOptionsMap,
            Map<String, List<AsSummaryTemplateTableColumn>> fieldColumnsMap) {
        
        // 按字段ID对所有字段建立索引
        Map<String, AsSummaryTemplateField> fieldMap = allFields.stream()
                .collect(Collectors.toMap(AsSummaryTemplateField::getId, field -> field));
        
        // 找出所有顶级字段（项目级别，parentFieldId为null的字段）
        List<AsSummaryTemplateField> projectFields = allFields.stream()
                .filter(field -> StringUtils.isBlank(field.getParentFieldId()))
                .sorted(Comparator.comparing(AsSummaryTemplateField::getSortOrder))
                .collect(Collectors.toList());
        
        // 将所有字段按照parentFieldId进行分组，用于快速查找子字段
        Map<String, List<AsSummaryTemplateField>> fieldsByParentId = allFields.stream()
                .filter(field -> StringUtils.isNotBlank(field.getParentFieldId()))
                .collect(Collectors.groupingBy(AsSummaryTemplateField::getParentFieldId));
        
        // 构建项目级节点
        return projectFields.stream().map(projectField -> {
            TemplateFieldNode projectNode = createNode(projectField, fieldOptionsMap, fieldColumnsMap, 0);
            
            // 获取该项目下的内容字段
            List<AsSummaryTemplateField> contentFields = fieldsByParentId.getOrDefault(projectField.getId(), Collections.emptyList())
                    .stream()
                    .sorted(Comparator.comparing(AsSummaryTemplateField::getSortOrder))
                    .collect(Collectors.toList());
            
            // 处理内容级节点
            contentFields.forEach(contentField -> {
                TemplateFieldNode contentNode = createNode(contentField, fieldOptionsMap, fieldColumnsMap, 1);
                
                // 获取该内容下的回复字段
                List<AsSummaryTemplateField> replyFields = fieldsByParentId.getOrDefault(contentField.getId(), Collections.emptyList())
                        .stream()
                        .sorted(Comparator.comparing(AsSummaryTemplateField::getSortOrder))
                        .collect(Collectors.toList());
                
                // 修改：直接将所有回复字段添加为节点，移除 label 判断逻辑
                replyFields.forEach(replyField -> {
                    TemplateFieldNode replyNode = createNode(replyField, fieldOptionsMap, fieldColumnsMap, 2);
                    contentNode.getChildren().add(replyNode);
                });
                
                projectNode.getChildren().add(contentNode);
            });
            
            return projectNode;
        }).collect(Collectors.toList());
    }
    
    /**
     * 创建节点并设置相关属性
     */
    private TemplateFieldNode createNode(
            AsSummaryTemplateField field,
            Map<String, List<AsSummaryTemplateOption>> fieldOptionsMap,
            Map<String, List<AsSummaryTemplateTableColumn>> fieldColumnsMap,
            int level) {
        
        TemplateFieldNode node = new TemplateFieldNode();
        node.setField(field);
        node.setLevel(level);
        node.setOptions(fieldOptionsMap.getOrDefault(field.getId(), Collections.emptyList()));
        node.setColumns(fieldColumnsMap.getOrDefault(field.getId(), Collections.emptyList()));
        return node;
    }
    
    /**
     * 新增：计算最大表格列数，递归遍历所有节点
     */
    private int getMaxTableCols(List<TemplateFieldNode> fieldTree) {
        int max = 0;
        for (TemplateFieldNode node : fieldTree) {
            if ("TABLE".equalsIgnoreCase(node.getField().getFieldType())) {
                max = Math.max(max, node.getColumns().size());
            }
            // 递归检查子节点
            max = Math.max(max, getMaxTableCols(node.getChildren()));
        }
        return max;
    }
    
    /**
     * 生成Excel文件
     *
     * @param templateName 模板名称
     * @param fieldTree 字段树状结构
     * @return Excel文件字节数组
     * @throws IOException 生成Excel异常
     */
    private byte[] generateExcel(String templateName, List<TemplateFieldNode> fieldTree) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            // 新增：计算最大表格列数和填表说明列索引
            int maxTableCols = getMaxTableCols(fieldTree);
            this.fillReasonCol = TABLE_START_COL + maxTableCols;

            Sheet sheet = workbook.createSheet(templateName);

            // 动态设置列宽
            sheet.setColumnWidth(0, 20 * 256);
            sheet.setColumnWidth(1, 25 * 256);
            for (int i = 0; i < maxTableCols; i++) {
                sheet.setColumnWidth(TABLE_START_COL + i, 20 * 256);
            }
            sheet.setColumnWidth(fillReasonCol, 25 * 256);

            // 创建样式
            CellStyle titleStyle = createTitleStyle(workbook);
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle contentStyle = createContentStyle(workbook);

            // 创建表头和标题行
            createHeaderAndTitle(sheet, templateName, titleStyle, headerStyle);

            // 写入所有数据并收集合并信息
            int dataStartRow = 2;
            List<MergeRegion> mergeRegions = new ArrayList<>();
            writeNodesData(sheet, fieldTree, dataStartRow, contentStyle, mergeRegions);

            // 执行单元格合并
            executeMerges(sheet, mergeRegions);

            // 输出Excel文件
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                workbook.write(outputStream);
                return outputStream.toByteArray();
            }
        }
    }
    
    /**
     * 重写表头和标题行，支持动态列合并
     */
    private void createHeaderAndTitle(Sheet sheet, String templateName, CellStyle titleStyle, CellStyle headerStyle) {
        // 模板名行
        Row headerRow = sheet.createRow(0);
        headerRow.setHeightInPoints(30);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue(templateName);
        headerCell.setCellStyle(headerStyle);
        // 合并模板名跨列
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, fillReasonCol));

        // 列标题行
        Row titleRow = sheet.createRow(1);
        titleRow.setHeightInPoints(25);
        // 项目列
        Cell cell0 = titleRow.createCell(0);
        cell0.setCellValue("项目");
        cell0.setCellStyle(headerStyle);
        // 内容列
        Cell cell1 = titleRow.createCell(1);
        cell1.setCellValue("内容");
        cell1.setCellStyle(headerStyle);
        // 回复组头，跨表格列
        Cell cellReply = titleRow.createCell(TABLE_START_COL);
        cellReply.setCellValue("回复");
        cellReply.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, TABLE_START_COL, fillReasonCol - 1));
        // 填表说明列
        Cell cellDesc = titleRow.createCell(fillReasonCol);
        cellDesc.setCellValue("填表说明");
        cellDesc.setCellStyle(headerStyle);
    }
    
    /**
     * 单元格合并区域信息
     */
    private static class MergeRegion {
        final int firstRow;
        final int lastRow;
        final int firstCol;
        final int lastCol;
        
        public MergeRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
            this.firstRow = firstRow;
            this.lastRow = lastRow;
            this.firstCol = firstCol;
            this.lastCol = lastCol;
        }
    }
    
    /**
     * 写入节点数据并收集合并区域
     */
    private int writeNodesData(Sheet sheet, List<TemplateFieldNode> nodes, int startRow,
                              CellStyle style, List<MergeRegion> mergeRegions) {
        int currentRow = startRow;
        
        for (TemplateFieldNode node : nodes) {
            currentRow = writeNodeWithChildren(sheet, node, currentRow, style, mergeRegions);
        }
        
        return currentRow;
    }
    
    /**
     * 写入节点及其子节点数据
     */
    private int writeNodeWithChildren(Sheet sheet, TemplateFieldNode node, int startRow,
                               CellStyle style, List<MergeRegion> mergeRegions) {
        int nodeStartRow = startRow;
        int currentRow = startRow;
        
        // 写入当前节点数据
        writeNodeContent(sheet, node, currentRow, style);
        
        // 处理子节点
        if (!node.getChildren().isEmpty()) {
            currentRow++;
            for (TemplateFieldNode child : node.getChildren()) {
                currentRow = writeNodeWithChildren(sheet, child, currentRow, style, mergeRegions);
            }
            if (currentRow - 1 > nodeStartRow) {
                int level = node.getLevel();
                if (level == 0) {
                    mergeRegions.add(new MergeRegion(nodeStartRow, currentRow - 1, 0, 0));
                } else if (level == 1) {
                    mergeRegions.add(new MergeRegion(nodeStartRow, currentRow - 1, 1, 1));
                }
            }
        } else {
            // 修改：TABLE 类型使用原生 Excel 列来渲染表格
            if ("TABLE".equalsIgnoreCase(node.getField().getFieldType()) && node.getLevel() == 2) {
                int tableRows = 5; // 固定5行（1行标题+4行数据）
                for (int i = 0; i < tableRows; i++) {
                    int rowIndex = nodeStartRow + i;
                    Row row = getOrCreateRow(sheet, rowIndex);
                    // 初始化整行样式
                    for (int c = 0; c <= fillReasonCol; c++) {
                        Cell cell = row.getCell(c);
                        if (cell == null) cell = row.createCell(c);
                        cell.setCellStyle(style);
                    }
                    // 写入表格列标题和空白行
                    for (int j = 0; j < node.getColumns().size(); j++) {
                        Cell cell = row.getCell(TABLE_START_COL + j);
                        cell.setCellStyle(style);
                        if (i == 0) {
                            cell.setCellValue(node.getColumns().get(j).getColumnLabel());
                        }
                    }
                    // 只有首行显示说明
                    if (i == 0 && StringUtils.isNotBlank(node.getField().getDescription())) {
                        row.getCell(fillReasonCol).setCellValue(node.getField().getDescription());
                    }
                }
                // 合并“填表说明”列单元格，使说明跨多行显示
                mergeRegions.add(new MergeRegion(nodeStartRow, nodeStartRow + tableRows - 1, fillReasonCol, fillReasonCol));
                currentRow = nodeStartRow + tableRows;
            } else {
                currentRow++;
            }
        }
        
        return currentRow;
    }
    
    /**
     * 创建空行（带边框）
     */
    private void createEmptyRow(Sheet sheet, int rowIndex, CellStyle style) {
        Row row = sheet.createRow(rowIndex);
        for (int i = 0; i < 4; i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(style);
        }
    }
    
    /**
     * 重写写入节点内容，支持动态描述列
     */
    private void writeNodeContent(Sheet sheet, TemplateFieldNode node, int rowIndex, CellStyle style) {
        Row row = getOrCreateRow(sheet, rowIndex);
        // 初始化单元格样式
        for (int c = 0; c <= fillReasonCol; c++) {
            Cell cell = row.getCell(c);
            if (cell == null) cell = row.createCell(c);
            cell.setCellStyle(style);
        }
        int level = node.getLevel();
        switch (level) {
            case 0:
                row.getCell(0).setCellValue(node.getField().getLabel());
                break;
            case 1:
                row.getCell(1).setCellValue(node.getField().getLabel());
                break;
            case 2:
                writeReplyContent(row, node);
                if (StringUtils.isNotBlank(node.getField().getDescription())) {
                    row.getCell(fillReasonCol).setCellValue(node.getField().getDescription());
                }
                break;
        }
    }
    
    /**
     * 获取或创建行
     */
    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }
    
    /**
     * 写入回复内容
     */
    private void writeReplyContent(Row row, TemplateFieldNode node) {
        Cell replyCell = row.getCell(2);
        String fieldType = node.getField().getFieldType();
        
        switch (fieldType) {
            case "SINGLE_CHOICE":
                // 单选，显示所有选项，逗号分隔
                if (!node.getOptions().isEmpty()) {
                    String optionsText = node.getOptions().stream()
                            .map(AsSummaryTemplateOption::getOptionLabel)
                            .collect(Collectors.joining("，"));
                    replyCell.setCellValue(optionsText);
                }
                break;
                
            case "TEXT":
                // 单行文本：前缀 + () + 后缀
                String prefix = StringUtils.defaultString(node.getField().getPrefix(), "");
                String suffix = StringUtils.defaultString(node.getField().getSuffix(), "");
                replyCell.setCellValue(prefix + "（     ）" + suffix);
                break;
                
            case "TABLE":
                // TABLE类型不在这里处理，留空即可，由writeNodeWithChildren中的特殊逻辑处理
                // 这里不要写任何内容，避免与TABLE的多行处理冲突
                break;
                
            case "TEXTAREA":
                // 文本域，留出填写空间
                replyCell.setCellValue("请在此处填写...");
                break;
        }
    }
    
    /**
     * 执行单元格合并，处理可能的重叠
     */
    private void executeMerges(Sheet sheet, List<MergeRegion> mergeRegions) {
        // 对合并区域按列和起始行排序
        mergeRegions.sort((a, b) -> a.firstCol != b.firstCol ? 
                a.firstCol - b.firstCol : a.firstRow - b.firstRow);
        
        // 跟踪每列已合并的区域
        Map<Integer, List<MergeRegion>> columnMerges = new HashMap<>();
        
        // 执行合并，确保不重叠
        for (MergeRegion region : mergeRegions) {
            // 检查是否有重叠
            boolean overlaps = false;
            List<MergeRegion> colMerges = columnMerges.computeIfAbsent(region.firstCol, k -> new ArrayList<>());
            
            for (MergeRegion existing : colMerges) {
                // 检查重叠条件
                if (!(region.lastRow < existing.firstRow || region.firstRow > existing.lastRow)) {
                    overlaps = true;
                    logger.warn("检测到重叠区域，跳过合并: [{},{}] 与 [{},{}]",
                            region.firstRow, region.lastRow, existing.firstRow, existing.lastRow);
                    break;
                }
            }
            
            if (!overlaps) {
                try {
                    sheet.addMergedRegion(new CellRangeAddress(
                            region.firstRow, region.lastRow, region.firstCol, region.lastCol));
                    colMerges.add(region);
                } catch (Exception e) {
                    logger.error("合并区域失败: [{},{}], 错误: {}", region.firstRow, region.lastRow, e.getMessage());
                }
            }
        }
    }
    
    /**
     * 创建标题样式（不带背景色）
     */
    private CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        
        return style;
    }
    
    /**
     * 创建表头样式（黄色背景）
     */
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        // 设置黄色背景
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        
        return style;
    }
    
    /**
     * 创建内容样式
     */
    private CellStyle createContentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true);
        return style;
    }
}