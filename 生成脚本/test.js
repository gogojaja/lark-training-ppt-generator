/**
 * test.js - 测试脚本
 * 用于测试组件功能
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

const ThemeManager = require('./components/ThemeManager');
const ComponentLoader = require('./components/ComponentLoader');
const SchemaValidator = require('./components/SchemaValidator');
const ConfigManager = require('./components/ConfigManager');
const SlideRenderer = require('./components/SlideRenderer');
const { parseArgs } = require('./generate.js');

function test() {
    console.log('开始组件测试...\n');
    const results = {
        passed: 0,
        failed: 0,
        errors: []
    };

    // 测试ThemeManager
    try {
        const themeManager = new ThemeManager();
        console.log('✓ ThemeManager 实例化成功');
        results.passed++;
        
        // 测试列出主题
        const themes = themeManager.listThemes();
        console.log(`✓ ThemeManager.listThemes() 返回 ${themes.length} 个主题`);
        results.passed++;
    } catch (error) {
        console.log('✗ ThemeManager 测试失败:', error.message);
        results.failed++;
        results.errors.push(`ThemeManager: ${error.message}`);
    }

    // 测试ComponentLoader
    try {
        const componentLoader = new ComponentLoader();
        console.log('✓ ComponentLoader 实例化成功');
        results.passed++;
        
        // 测试列出组件
        const components = componentLoader.listComponents();
        console.log(`✓ ComponentLoader.listComponents() 返回 ${components.length} 个组件`);
        results.passed++;
    } catch (error) {
        console.log('✗ ComponentLoader 测试失败:', error.message);
        results.failed++;
        results.errors.push(`ComponentLoader: ${error.message}`);
    }

    // 测试SchemaValidator
    try {
        const schemaValidator = new SchemaValidator();
        console.log('✓ SchemaValidator 实例化成功');
        results.passed++;
    } catch (error) {
        console.log('✗ SchemaValidator 测试失败:', error.message);
        results.failed++;
        results.errors.push(`SchemaValidator: ${error.message}`);
    }

    // 测试ConfigManager
    try {
        const configManager = new ConfigManager();
        console.log('✓ ConfigManager 实例化成功');
        results.passed++;
        
        // 测试获取配置
        const config = configManager.getConfig();
        console.log('✓ ConfigManager.getConfig() 成功');
        results.passed++;
    } catch (error) {
        console.log('✗ ConfigManager 测试失败:', error.message);
        results.failed++;
        results.errors.push(`ConfigManager: ${error.message}`);
    }

    // 测试SlideRenderer
    try {
        const componentLoader = new ComponentLoader();
        const slideRenderer = new SlideRenderer(componentLoader);
        console.log('✓ SlideRenderer 实例化成功');
        results.passed++;
    } catch (error) {
        console.log('✗ SlideRenderer 测试失败:', error.message);
        results.failed++;
        results.errors.push(`SlideRenderer: ${error.message}`);
    }

    // 测试CoverComponent
    try {
        const componentLoader = new ComponentLoader();
        const coverComponent = componentLoader.loadComponent('CoverComponent');
        const testSlide = {
            page: 1,
            type: 'cover',
            key_message: '测试封面',
            content_elements: {
                title: '测试标题',
                subtitle: '测试副标题'
            }
        };
        const testTheme = {
            palette: { primary: '065A82', accent: 'E86A33' },
            fonts: { heading: 'Microsoft YaHei', body: 'Microsoft YaHei' }
        };
        const result = coverComponent.render(testSlide, testTheme);
        console.log('✓ CoverComponent.render() 成功');
        results.passed++;
    } catch (error) {
        console.log('✗ CoverComponent 测试失败:', error.message);
        results.failed++;
        results.errors.push(`CoverComponent: ${error.message}`);
    }

    // 测试CLI参数解析
    try {
        const args = parseArgs(['--theme', 'charcoal-minimal', '--output', 'demo.pptx', '--local-only', '--export', 'lark']);
        if (args.theme !== 'charcoal-minimal' || args.output !== 'demo.pptx' || args.localOnly !== true || args.export !== 'lark') {
            throw new Error('CLI参数解析异常');
        }
        console.log('✓ parseArgs() 解析成功');
        results.passed++;
    } catch (error) {
        console.log('✗ parseArgs() 测试失败:', error.message);
        results.failed++;
        results.errors.push(`parseArgs: ${error.message}`);
    }

    // 测试配置驱动主题解析
    try {
        const { resolveEffectiveConfig } = require('./generate.js');
        const config = resolveEffectiveConfig(['--config', 'config.yaml', '--theme', 'charcoal-minimal']);
        if (config.theme !== 'charcoal-minimal') {
            throw new Error('配置和CLI主题未按优先级生效');
        }
        console.log('✓ resolveEffectiveConfig() 生效');
        results.passed++;
    } catch (error) {
        console.log('✗ resolveEffectiveConfig() 测试失败:', error.message);
        results.failed++;
        results.errors.push(`resolveEffectiveConfig: ${error.message}`);
    }

    // 总结
    console.log('\n测试总结:');
    console.log(`通过: ${results.passed}`);
    console.log(`失败: ${results.failed}`);
    console.log(`错误: ${results.errors.length}`);
    
    if (results.errors.length > 0) {
        console.log('\n错误详情:');
        results.errors.forEach(error => console.log(`  - ${error}`));
    }

    return results.failed === 0;
}

if (require.main === module) {
    const success = test();
    process.exit(success ? 0 : 1);
}

module.exports = test;
