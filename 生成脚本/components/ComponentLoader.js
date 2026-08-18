/**
 * ComponentLoader - 组件加载器
 * 负责动态加载components/目录下的组件
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

const fs = require('fs');
const path = require('path');

class ComponentLoader {
    constructor(componentsDir) {
        this.componentsDir = componentsDir || path.join(__dirname);
        this.loadedComponents = new Map();
    }

    /**
     * 加载组件
     * @param {string} componentType - 组件类型名称
     * @returns {Object} 组件实例
     */
    loadComponent(componentType) {
        if (this.loadedComponents.has(componentType)) {
            return this.loadedComponents.get(componentType);
        }

        const componentPath = path.join(this.componentsDir, `${componentType}.js`);
        
        if (!fs.existsSync(componentPath)) {
            throw new Error(`组件文件不存在: ${componentPath}`);
        }

        try {
            const ComponentClass = require(componentPath);
            const component = new ComponentClass();
            this.loadedComponents.set(componentType, component);
            return component;
        } catch (error) {
            throw new Error(`加载组件失败 ${componentType}: ${error.message}`);
        }
    }

    /**
     * 列出所有可用组件
     * @returns {Array} 组件类型列表
     */
    listComponents() {
        const components = [];
        if (fs.existsSync(this.componentsDir)) {
            const files = fs.readdirSync(this.componentsDir);
            for (const file of files) {
                if (file.endsWith('.js') && file !== 'ComponentLoader.js') {
                    components.push(path.basename(file, '.js'));
                }
            }
        }
        return components;
    }

    /**
     * 卸载组件
     * @param {string} componentType - 组件类型名称
     */
    unloadComponent(componentType) {
        if (this.loadedComponents.has(componentType)) {
            this.loadedComponents.delete(componentType);
        }
    }

    /**
     * 卸载所有组件
     */
    unloadAllComponents() {
        this.loadedComponents.clear();
    }

    /**
     * 检查组件是否存在
     * @param {string} componentType - 组件类型名称
     * @returns {boolean} 是否存在
     */
    hasComponent(componentType) {
        const componentPath = path.join(this.componentsDir, `${componentType}.js`);
        return fs.existsSync(componentPath);
    }
}

module.exports = ComponentLoader;
