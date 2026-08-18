/**
 * SchemaValidator - Schema验证器
 * 负责验证JSON/YAML文件格式
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

const fs = require('fs');
const path = require('path');

class SchemaValidator {
    constructor(schemasDir) {
        this.schemasDir = schemasDir || path.join(__dirname, '..', 'schemas');
        this.schemasCache = new Map();
    }

    /**
     * 加载Schema文件
     * @param {string} schemaName - Schema名称
     * @returns {Object} Schema对象
     */
    loadSchema(schemaName) {
        if (this.schemasCache.has(schemaName)) {
            return this.schemasCache.get(schemaName);
        }

        const schemaPath = path.join(this.schemasDir, `${schemaName}.json`);
        
        if (!fs.existsSync(schemaPath)) {
            throw new Error(`Schema文件不存在: ${schemaPath}`);
        }

        const schema = JSON.parse(fs.readFileSync(schemaPath, 'utf-8'));
        this.schemasCache.set(schemaName, schema);
        return schema;
    }

    /**
     * 验证数据是否符合Schema
     * @param {Object} schema - Schema对象
     * @param {Object} data - 待验证数据
     * @returns {Object} 验证结果 {valid: boolean, errors: Array}
     */
    validate(schema, data) {
        const errors = [];
        
        // 简化的验证逻辑
        if (schema.required) {
            for (const field of schema.required) {
                if (!(field in data)) {
                    errors.push(`缺少必需字段: ${field}`);
                }
            }
        }

        if (schema.properties) {
            for (const [key, propSchema] of Object.entries(schema.properties)) {
                if (key in data) {
                    const value = data[key];
                    if (propSchema.type) {
                        if (propSchema.type === 'string' && typeof value !== 'string') {
                            errors.push(`字段 ${key} 应为字符串类型`);
                        } else if (propSchema.type === 'number' && typeof value !== 'number') {
                            errors.push(`字段 ${key} 应为数字类型`);
                        } else if (propSchema.type === 'array' && !Array.isArray(value)) {
                            errors.push(`字段 ${key} 应为数组类型`);
                        } else if (propSchema.type === 'object' && typeof value !== 'object') {
                            errors.push(`字段 ${key} 应为对象类型`);
                        }
                    }
                    if (propSchema.properties && typeof value === 'object') {
                        const nestedResult = this.validate(propSchema, value);
                        if (!nestedResult.valid) {
                            errors.push(...nestedResult.errors.map(e => `${key}.${e}`));
                        }
                    }
                }
            }
        }

        return {
            valid: errors.length === 0,
            errors
        };
    }

    /**
     * 验证slide_plan.json
     * @param {Object} slidePlan - slide_plan数据
     * @returns {Object} 验证结果
     */
    validateSlidePlan(slidePlan) {
        const schema = this.loadSchema('slide_plan');
        return this.validate(schema, slidePlan);
    }

    /**
     * 验证主题配置
     * @param {Object} theme - 主题配置
     * @returns {Object} 验证结果
     */
    validateTheme(theme) {
        const schema = this.loadSchema('theme');
        return this.validate(schema, theme);
    }

    /**
     * 验证config.yaml解析后的对象
     * @param {Object} config - 配置对象
     * @returns {Object} 验证结果
     */
    validateConfig(config) {
        const schema = this.loadSchema('config');
        return this.validate(schema, config);
    }
}

module.exports = SchemaValidator;
