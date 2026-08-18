/**
 * CoverComponent - 封面组件
 * 负责渲染封面幻灯片
 * 
 * @version 1.0.0
 * @date 2026-08-07
 */

class CoverComponent {
    /**
     * 渲染封面幻灯片
     * @param {Object} slide - slide配置
     * @param {Object} theme - 主题配置
     * @returns {Object} 渲染结果
     */
    render(slide, theme) {
        const elements = [];
        
        // 背景
        elements.push({
            type: 'background',
            color: theme.palette.dark || '21295C'
        });
        
        // 主标题
        elements.push({
            type: 'text',
            content: slide.content_elements?.title || slide.key_message,
            style: {
                fontSize: 40,
                fontFace: theme.fonts.heading,
                color: theme.palette.white || 'FFFFFF',
                bold: true,
                align: 'center',
                valign: 'middle',
                x: 0.8,
                y: 1.5,
                w: 8.4,
                h: 0.8
            }
        });
        
        // 副标题
        elements.push({
            type: 'text',
            content: slide.content_elements?.subtitle || '',
            style: {
                fontSize: 22,
                fontFace: theme.fonts.heading,
                color: 'B0D4E8',
                align: 'center',
                valign: 'middle',
                x: 0.8,
                y: 2.35,
                w: 8.4,
                h: 0.5
            }
        });
        
        // 强调线
        elements.push({
            type: 'shape',
            shape: 'rectangle',
            style: {
                x: 3.5,
                y: 3.0,
                w: 3,
                h: 0.03,
                fill: { color: theme.palette.accent }
            }
        });
        
        // 受众信息
        elements.push({
            type: 'text',
            content: slide.content_elements?.audience_label || '',
            style: {
                fontSize: 13,
                fontFace: theme.fonts.body,
                color: 'D1D5DB',
                align: 'center',
                valign: 'middle',
                x: 0.8,
                y: 3.2,
                w: 8.4,
                h: 0.35
            }
        });
        
        // 交易信息
        if (slide.content_elements?.transaction_code) {
            elements.push({
                type: 'text',
                content: `交易代码 ${slide.content_elements.transaction_code} | 权限：${slide.content_elements.permission || ''}`,
                style: {
                    fontSize: 11,
                    fontFace: theme.fonts.body,
                    color: '9CA3AF',
                    align: 'center',
                    valign: 'middle',
                    x: 0.8,
                    y: 4.2,
                    w: 8.4,
                    h: 0.3
                }
            });
        }
        
        return {
            type: 'cover',
            elements,
            notes: slide.notice_items?.join('\n') || ''
        };
    }
}

module.exports = CoverComponent;
