/**
 * ConfigManager - 配置管理器
 * 负责读取和管理配置文件
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

const fs = require('fs');
const path = require('path');

class ConfigManager {
    constructor(configPath) {
        this.configPath = configPath || path.join(__dirname, '..', 'config.yaml');
        this.config = null;
        this.defaultConfig = {
            theme: 'ocean-gradient',
            export_format: 'pptx',
            page_size: '16:9',
            brand_guideline: null,
            output_dir: '生成产物'
        };
    }

    /**
     * 读取配置文件
     * @returns {Object} 配置对象
     */
    getConfig() {
        if (this.config) {
            return this.config;
        }

        if (fs.existsSync(this.configPath)) {
            const configContent = fs.readFileSync(this.configPath, 'utf-8');
            this.config = this.parseYAML(configContent);
        } else {
            this.config = { ...this.defaultConfig };
            this.saveConfig();
        }

        return this.config;
    }

    /**
     * 简化的YAML解析器
     * @param {string} yamlContent - YAML内容
     * @returns {Object} 解析后的对象
     */
    parseYAML(yamlContent) {
        const result = {};
        const lines = yamlContent.split('\n');
        
        for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed || trimmed.startsWith('#')) {
                continue;
            }
            
            const colonIndex = trimmed.indexOf(':');
            if (colonIndex > 0) {
                const key = trimmed.substring(0, colonIndex).trim();
                let value = trimmed.substring(colonIndex + 1).trim();
                
                // 处理引号
                if ((value.startsWith('"') && value.endsWith('"')) ||
                    (value.startsWith("'") && value.endsWith("'"))) {
                    value = value.slice(1, -1);
                } else if (value === 'true') {
                    value = true;
                } else if (value === 'false') {
                    value = false;
                } else if (value === 'null' || value === '~') {
                    value = null;
                } else if (!isNaN(value) && value !== '') {
                    value = Number(value);
                }
                
                result[key] = value;
            }
        }
        
        return result;
    }

    /**
     * 简化的YAML序列化器
     * @param {Object} obj - 对象
     * @returns {string} YAML字符串
     */
    stringifyYAML(obj) {
        const lines = [];
        for (const [key, value] of Object.entries(obj)) {
            if (value === null || value === undefined) {
                lines.push(`${key}: null`);
            } else if (typeof value === 'string') {
                lines.push(`${key}: "${value}"`);
            } else if (typeof value === 'boolean') {
                lines.push(`${key}: ${value}`);
            } else if (typeof value === 'number') {
                lines.push(`${key}: ${value}`);
            } else {
                lines.push(`${key}: "${JSON.stringify(value)}"`);
            }
        }
        return lines.join('\n');
    }

    /**
     * 保存配置文件
     */
    saveConfig() {
        if (!this.config) {
            this.config = { ...this.defaultConfig };
        }
        
        const yamlContent = this.stringifyYAML(this.config);
        fs.writeFileSync(this.configPath, yamlContent);
    }

    /**
     * 更新配置
     * @param {Object} newConfig - 新配置
     */
    updateConfig(newConfig) {
        this.config = { ...this.getConfig(), ...newConfig };
        this.saveConfig();
    }

    /**
     * 获取默认配置
     * @returns {Object} 默认配置
     */
    getDefaultConfig() {
        return { ...this.defaultConfig };
    }

    /**
     * 重置为默认配置
     */
    resetToDefault() {
        this.config = { ...this.defaultConfig };
        this.saveConfig();
    }
}

module.exports = ConfigManager;
