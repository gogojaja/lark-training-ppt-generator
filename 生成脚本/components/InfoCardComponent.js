/**
 * InfoCardComponent - 信息卡片组件
 * 负责渲染信息卡片类型的幻灯片
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

class InfoCardComponent {
    /**
     * 渲染信息卡片幻灯片
     * @param {Object} slide - slide配置
     * @param {Object} theme - 主题配置
     * @returns {Object} 渲染结果
     */
    render(slide, theme) {
        const elements = [];
        
        // 页眉
        elements.push(this.renderHeader(slide, theme));
        
        // 内容区域
        if (slide.content_elements) {
            // 场景定义卡片
            if (slide.content_elements.scene_definition) {
                elements.push(this.renderCard(
                    slide.content_elements.scene_definition,
                    theme.palette.primary,
                    1.05, 4.3, 2.0
                ));
            }
            
            // 账户类型卡片
            if (slide.content_elements.account_types) {
                elements.push(this.renderCard(
                    slide.content_elements.account_types,
                    theme.palette.secondary,
                    5.55, 1.05, 4.3, 2.0
                ));
            }
            
            // 核心要点栏
            if (slide.content_elements.key_points) {
                elements.push(this.renderKeyPoints(
                    slide.content_elements.key_points,
                    theme
                ));
            }
        }
        
        // 页脚
        elements.push(this.renderFooter(slide, theme));
        
        return {
            type: 'info_card',
            elements,
            notes: slide.notice_items?.join('\n') || ''
        };
    }

    renderHeader(slide, theme) {
        return {
            type: 'header',
            page: slide.page,
            title: slide.key_message,
            subtitle: slide.layout_type,
            theme
        };
    }

    renderCard(content, borderColor, x, y, w, h) {
        return {
            type: 'card',
            content,
            borderColor,
            x, y, w, h
        };
    }

    renderKeyPoints(points, theme) {
        return {
            type: 'key_points',
            points,
            bgColor: theme.palette.dark
        };
    }

    renderFooter(slide, theme) {
        return {
            type: 'footer',
            text: '培训内容'
        };
    }
}

module.exports = InfoCardComponent;
