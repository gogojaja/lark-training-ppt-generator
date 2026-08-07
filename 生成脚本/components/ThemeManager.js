/**
 * ThemeManager - 主题管理器
 * 负责主题的加载、切换、验证
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

const fs = require('fs');
const path = require('path');

class ThemeManager {
    constructor(themesDir) {
        this.themesDir = themesDir || path.join(__dirname, '..', 'themes');
        this.currentTheme = null;
        this.themesCache = new Map();
    }

    /**
     * 加载主题配置
     * @param {string} themeName - 主题名称或路径
     * @returns {Object} 主题配置对象
     */
    getTheme(themeName) {
        if (this.themesCache.has(themeName)) {
            return this.themesCache.get(themeName);
        }

        let themePath;
        if (path.isAbsolute(themeName)) {
            themePath = themeName;
        } else {
            themePath = path.join(this.themesDir, `${themeName}.json`);
        }

        if (!fs.existsSync(themePath)) {
            throw new Error(`主题文件不存在: ${themePath}`);
        }

        const themeData = JSON.parse(fs.readFileSync(themePath, 'utf-8'));
        const validatedTheme = this.validateTheme(themeData);
        
        this.themesCache.set(themeName, validatedTheme);
        return validatedTheme;
    }

    /**
     * 验证主题配置格式
     * @param {Object} theme - 主题配置对象
     * @returns {Object} 验证后的主题配置
     */
    validateTheme(theme) {
        const requiredFields = ['name', 'palette', 'fonts'];
        const requiredPalette = ['primary', 'secondary', 'accent', 'bg', 'text'];
        const requiredFonts = ['heading', 'body'];

        for (const field of requiredFields) {
            if (!theme[field]) {
                throw new Error(`主题缺少必需字段: ${field}`);
            }
        }

        for (const field of requiredPalette) {
            if (!theme.palette[field]) {
                throw new Error(`主题palette缺少必需字段: ${field}`);
            }
        }

        for (const field of requiredFonts) {
            if (!theme.fonts[field]) {
                throw new Error(`主题fonts缺少必需字段: ${field}`);
            }
        }

        return theme;
    }

    /**
     * 设置当前活动主题
     * @param {string} themeName - 主题名称
     */
    setCurrentTheme(themeName) {
        const theme = this.getTheme(themeName);
        this.currentTheme = theme;
        return theme;
    }

    /**
     * 获取当前活动主题
     * @returns {Object} 当前主题配置
     */
    getCurrentTheme() {
        if (!this.currentTheme) {
            return this.getTheme('ocean-gradient');
        }
        return this.currentTheme;
    }

    /**
     * 列出所有可用主题
     * @returns {Array} 主题列表
     */
    listThemes() {
        const themes = [];
        if (fs.existsSync(this.themesDir)) {
            const files = fs.readdirSync(this.themesDir);
            for (const file of files) {
                if (file.endsWith('.json')) {
                    themes.push(path.basename(file, '.json'));
                }
            }
        }
        return themes;
    }

    /**
     * 创建默认主题
     * @param {string} name - 主题名称
     * @returns {Object} 创建的主题配置
     */
    createDefaultTheme(name) {
        const defaultTheme = {
            name: name,
            palette: {
                primary: '065A82',
                secondary: '1C7293',
                accent: 'E86A33',
                bg: 'F5F7FA',
                text: '1F2937',
                textLight: '6B7280',
                success: '2D9C5E',
                warning: 'E8A838',
                danger: 'DC2626'
            },
            fonts: {
                heading: 'Microsoft YaHei',
                body: 'Microsoft YaHei'
            },
            sizes: {
                title: '20pt',
                body: '9-10pt',
                note: '8pt'
            }
        };

        const themePath = path.join(this.themesDir, `${name}.json`);
        fs.writeFileSync(themePath, JSON.stringify(defaultTheme, null, 2));
        this.themesCache.set(name, defaultTheme);
        return defaultTheme;
    }
}

module.exports = ThemeManager;
