/**
 * SlideRenderer - 幻灯片渲染器
 * 负责遍历slide_plan调用对应组件渲染
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

const fs = require('fs');
const path = require('path');

class SlideRenderer {
    constructor(componentLoader) {
        this.componentLoader = componentLoader;
        this.renderedSlides = [];
    }

    /**
     * 渲染单个幻灯片
     * @param {Object} slide - slide_plan中的单页配置
     * @param {Object} theme - 主题配置
     * @returns {Object} 渲染结果
     */
    renderSlide(slide, theme) {
        const componentType = this.mapSlideTypeToComponent(slide.type);
        
        try {
            const component = this.componentLoader.loadComponent(componentType);
            const rendered = component.render(slide, theme);
            this.renderedSlides.push({
                page: slide.page,
                type: slide.type,
                component: componentType,
                content: rendered
            });
            return rendered;
        } catch (error) {
            console.error(`渲染幻灯片 ${slide.page} 失败: ${error.message}`);
            return this.renderFallback(slide, theme);
        }
    }

    /**
     * 将slide type映射到组件类型
     * @param {string} slideType - slide类型
     * @returns {string} 组件类型
     */
    mapSlideTypeToComponent(slideType) {
        const mapping = {
            'cover': 'CoverComponent',
            'toc': 'TocComponent',
            'scene_description': 'InfoCardComponent',
            'channel_comparison': 'TableComponent',
            'process_overview': 'StepCardsComponent',
            'flowchart': 'FlowchartComponent',
            'operation_overview': 'StepCardsComponent',
            'step_detail': 'StepDetailComponent',
            'step_precautions': 'NoticeComponent',
            'special_scenarios': 'RiskCardComponent',
            'quick_reference_closing': 'ReferenceTableComponent'
        };
        return mapping[slideType] || 'InfoCardComponent';
    }

    /**
     * 渲染所有幻灯片
     * @param {Array} slides - slide数组
     * @param {Object} theme - 主题配置
     * @returns {Array} 渲染结果数组
     */
    renderAll(slides, theme) {
        this.renderedSlides = [];
        
        for (const slide of slides) {
            this.renderSlide(slide, theme);
        }
        
        return this.renderedSlides;
    }

    /**
     * 兜底渲染
     * @param {Object} slide - slide配置
     * @param {Object} theme - 主题配置
     * @returns {Object} 兜底渲染结果
     */
    renderFallback(slide, theme) {
        return {
            type: 'fallback',
            title: slide.key_message || '内容',
            content: '组件渲染失败，使用兜底方案'
        };
    }

    /**
     * 获取渲染统计
     * @returns {Object} 统计信息
     */
    getStats() {
        const stats = {
            total: this.renderedSlides.length,
            byType: {}
        };

        for (const slide of this.renderedSlides) {
            stats.byType[slide.type] = (stats.byType[slide.type] || 0) + 1;
        }

        return stats;
    }

    /**
     * 清空渲染结果
     */
    clear() {
        this.renderedSlides = [];
    }
}

module.exports = SlideRenderer;
