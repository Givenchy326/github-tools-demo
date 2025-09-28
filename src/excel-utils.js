/**
 * Excel处理工具类
 * 提供JSON到Excel转换等实用功能
 */

/**
 * 将JSON数据转换为Excel格式的配置对象
 * @param {Array} data - 要转换的JSON数据数组
 * @param {string} filename - 生成的Excel文件名
 * @returns {Object} Excel配置对象
 */
function createExcelConfig(data, filename = 'data.xlsx') {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('数据必须是非空数组');
  }

  // 获取所有键作为表头
  const headers = Object.keys(data[0]);
  
  // 创建表头行
  const headerRow = headers.map(header => ({
    value: header,
    type: 'string'
  }));

  // 创建数据行
  const dataRows = data.map(row => 
    headers.map(header => ({
      value: row[header] || '',
      type: typeof row[header] === 'number' ? 'number' : 'string'
    }))
  );

  return {
    filename,
    sheet: {
      data: [headerRow, ...dataRows]
    }
  };
}

/**
 * 验证Excel数据格式
 * @param {Object} config - Excel配置对象
 * @returns {boolean} 是否为有效格式
 */
function validateExcelConfig(config) {
  if (!config || typeof config !== 'object') {
    return false;
  }

  if (!config.filename || typeof config.filename !== 'string') {
    return false;
  }

  if (!config.sheet || !Array.isArray(config.sheet.data)) {
    return false;
  }

  return true;
}

/**
 * 格式化数据用于Excel显示
 * @param {*} value - 原始值
 * @param {string} type - 数据类型
 * @returns {*} 格式化后的值
 */
function formatCellValue(value, type = 'string') {
  switch (type) {
    case 'number':
      return Number(value) || 0;
    case 'date':
      return new Date(value).toISOString().split('T')[0];
    case 'boolean':
      return Boolean(value) ? '是' : '否';
    default:
      return String(value || '');
  }
}

module.exports = {
  createExcelConfig,
  validateExcelConfig,
  formatCellValue
};