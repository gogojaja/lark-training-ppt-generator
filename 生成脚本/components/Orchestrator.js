/**
 * Orchestrator - 主控制器
 * 协调配置/主题/组件/输出的主控制器
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

const fs = require('fs');
const path = require('path');
const ThemeManager = require('./ThemeManager');
const ComponentLoader = require('./ComponentLoader');
const SchemaValidator = require('./SchemaValidator');
const ConfigManager = require('./ConfigManager');
const SlideRenderer = require('./SlideRenderer');

class Orchestrator {
    constructor(options = {}) {
        this.baseDir = options.baseDir || path.join(__dirname, '..');
        this.themeManager = new ThemeManager(path.join(this.baseDir, 'themes'));
        this.componentLoader = new ComponentLoader(path.join(__dirname));
        this.schemaValidator = new SchemaValidator(path.join(this.baseDir, 'schemas'));
        this.configManager = new ConfigManager(options.configPath);
        this.slideRenderer = new SlideRenderer(this.componentLoader);
    }

    /**
     * 生成PPT
     * @param {Object} options - 生成选项
     * @returns {Object} 生成结果
     */
    async generate(options = {}) {
        const startTime = Date.now();
        
        try {
            // 1. 加载配置
            const config = this.configManager.getConfig();
            const themeName = options.theme || config.theme;
            
            // 2. 加载主题
            const theme = this.themeManager.getTheme(themeName);
            
            // 3. 加载slide_plan
            const slidePlanPath = options.slidePlan || this.findSlidePlan(config);
            const slidePlan = this.loadSlidePlan(slidePlanPath);
            
            // 4. 验证slide_plan
            const validation = this.schemaValidator.validateSlidePlan(slidePlan);
            if (!validation.valid) {
                throw new Error(`slide_plan验证失败: ${validation.errors.join(', ')}`);
            }
            
            // 5. 渲染幻灯片
            const renderedSlides = this.slideRenderer.renderAll(slidePlan.slides, theme);
            
            // 6. 生成PPTX
            const pptxContent = this.buildPPTX(renderedSlides, slidePlan);
            
            // 7. 保存文件
            const outputPath = this.getOutputPath(options, config);
            this.savePPTX(pptxContent, outputPath);
            
            const endTime = Date.now();
            
            return {
                success: true,
                outputPath,
                stats: {
                    slides: renderedSlides.length,
                    theme: themeName,
                    duration: endTime - startTime
                }
            };
            
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * 查找slide_plan文件
     * @param {Object} config - 配置对象
     * @returns {string} slide_plan路径
     */
    findSlidePlan(config) {
        const possiblePaths = [
            path.join(this.baseDir, 'slide_plan.json'),
            path.join(this.baseDir, '生成脚本', 'slide_plan.json')
        ];

        for (const p of possiblePaths) {
            if (fs.existsSync(p)) {
                return p;
            }
        }

        throw new Error('未找到slide_plan.json文件');
    }

    /**
     * 加载slide_plan
     * @param {string} slidePlanPath - slide_plan路径
     * @returns {Object} slide_plan对象
     */
    loadSlidePlan(slidePlanPath) {
        const content = fs.readFileSync(slidePlanPath, 'utf-8');
        return JSON.parse(content);
    }

    /**
     * 构建PPTX
     * @param {Array} renderedSlides - 渲染后的幻灯片
     * @param {Object} slidePlan - slide_plan对象
     * @returns {Buffer} PPTX内容
     */
    buildPPTX(renderedSlides, slidePlan) {
        // 这里应该调用PPTXBuilder组件
        // 简化实现：返回渲染结果的JSON
        return JSON.stringify({
            meta: {
                title: slidePlan.title,
                totalSlides: renderedSlides.length
            },
            slides: renderedSlides
        });
    }

    /**
     * 获取输出路径
     * @param {Object} options - 选项
     * @param {Object} config - 配置
     * @returns {string} 输出路径
     */
    getOutputPath(options, config) {
        if (options.output) {
            return options.output;
        }

        const outputDir = path.join(this.baseDir, config.output_dir || '生成产物');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const fileName = `${slidePlan?.title || '生成结果'}.pptx`;
        return path.join(outputDir, fileName);
    }

    /**
     * 保存PPTX
     * @param {Buffer} content - PPTX内容
     * @param {string} outputPath - 输出路径
     */
    savePPTX(content, outputPath) {
        fs.writeFileSync(outputPath, content);
        console.log(`已生成: ${outputPath}`);
    }

    /**
     * 列出可用方案
     * @returns {Array} 方案列表
     */
    listSchemes() {
        const schemes = [];
        const files = fs.readdirSync(this.baseDir);
        
        for (const file of files) {
            if (file.startsWith('slide_plan') && file.endsWith('.json')) {
                schemes.push({
                    name: path.basename(file, '.json'),
                    path: path.join(this.baseDir, file)
                });
            }
        }
        
        return schemes;
    }
}

module.exports = Orchestrator;
