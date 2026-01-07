// ==========================================
// 核心配置逻辑：DataMappingConfig
// ==========================================

const handleConfigUpdate = (newConfig) => {
    // 1. 获取当前锁定的结果表信息 (Snapshot)
    const originalResultTable = currentConfig.resultTable;

    // 2. 强制覆盖结果表信息，确保不可被修改
    // 无论前端传入什么新的 ResultTable 配置，都强制重置为原有配置
    const finalConfig = {
        ...newConfig,
        resultTable: {
            ...originalResultTable, // ⚡️ 核心逻辑：结果表信息回滚至初始状态
            // 如果有特定的只读属性需要确保不被覆盖，也可以在这里显式声明
            tableName: originalResultTable.tableName,
            schema: originalResultTable.schema
        },
        // 3. 允许修改源表和映射
        sourceTable: newConfig.sourceTable, // 允许变动
        fieldMapping: newConfig.fieldMapping // 允许变动
    };

    // 4. 执行更新
    saveConfiguration(finalConfig);
};
