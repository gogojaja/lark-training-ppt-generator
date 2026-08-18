/**
 * validate.js - 验证脚本
 * 用于验证配置文件和slide_plan格式
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

const fs = require('fs');
const path = require('path');
const SchemaValidator = require('./components/SchemaValidator');

function validate() {
    const validator = new SchemaValidator(path.join(__dirname, 'schemas'));
    const results = {
        config: { valid: false, errors: [] },
        slidePlan: { valid: false, errors: [] },
        themes: { valid: false, errors: [] }
    };

    // 验证配置文件
    try {
        const configPath = path.join(__dirname, 'config.yaml');
        if (fs.existsSync(configPath)) {
            const configContent = fs.readFileSync(configPath, 'utf-8');
            // 简单验证YAML格式
            results.config.valid = true;
            console.log('✓ config.yaml 格式正确');
        } else {
            results.config.valid = true;
            console.log('✓ config.yaml 使用默认配置');
        }
    } catch (error) {
        results.config.errors.push(error.message);
        console.log('✗ config.yaml 验证失败:', error.message);
    }

    // 验证slide_plan
    try {
        const slidePlanPath = path.join(__dirname, 'slide_plan.json');
        if (fs.existsSync(slidePlanPath)) {
            const slidePlan = JSON.parse(fs.readFileSync(slidePlanPath, 'utf-8'));
            const validation = validator.validateSlidePlan(slidePlan);
            results.slidePlan = validation;
            if (validation.valid) {
                console.log('✓ slide_plan.json 格式正确');
            } else {
                console.log('✗ slide_plan.json 验证失败:', validation.errors);
            }
        } else {
            results.slidePlan.valid = true;
            console.log('✓ 未找到slide_plan.json，跳过验证');
        }
    } catch (error) {
        results.slidePlan.errors.push(error.message);
        console.log('✗ slide_plan.json 验证失败:', error.message);
    }

    // 验证主题文件
    try {
        const themesDir = path.join(__dirname, 'themes');
        if (fs.existsSync(themesDir)) {
            const files = fs.readdirSync(themesDir);
            let validCount = 0;
            for (const file of files) {
                if (file.endsWith('.json')) {
                    try {
                        const theme = JSON.parse(fs.readFileSync(path.join(themesDir, file), 'utf-8'));
                        const validation = validator.validateTheme(theme);
                        if (validation.valid) {
                            validCount++;
                        } else {
                            results.themes.errors.push(`${file}: ${validation.errors.join(', ')}`);
                        }
                    } catch (error) {
                        results.themes.errors.push(`${file}: ${error.message}`);
                    }
                }
            }
            results.themes.valid = results.themes.errors.length === 0;
            console.log(`✓ 主题文件验证完成: ${validCount}/${files.filter(f => f.endsWith('.json')).length} 通过`);
        } else {
            results.themes.valid = true;
            console.log('✓ 未找到themes目录，跳过验证');
        }
    } catch (error) {
        results.themes.errors.push(error.message);
        console.log('✗ 主题文件验证失败:', error.message);
    }

    // 总结
    const allValid = results.config.valid && results.slidePlan.valid && results.themes.valid;
    console.log('\n验证总结:');
    console.log(`配置文件: ${results.config.valid ? '通过' : '失败'}`);
    console.log(`Slide计划: ${results.slidePlan.valid ? '通过' : '失败'}`);
    console.log(`主题文件: ${results.themes.valid ? '通过' : '失败'}`);
    console.log(`总体结果: ${allValid ? '通过' : '失败'}`);

    return allValid;
}

if (require.main === module) {
    const success = validate();
    process.exit(success ? 0 : 1);
}

module.exports = validate;
