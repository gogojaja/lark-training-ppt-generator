/* 各场景制作耗时对比 - ECharts 分组柱状图
 * 数据来源：handbook-deck-overview.html "提效场景示例" 表格
 * 耗时统一换算为分钟（1 个工作日 = 480 分钟）
 */
(function () {
  'use strict';

  function ready(fn) {
    if (document.readyState !== 'loading') {
      fn();
    } else {
      document.addEventListener('DOMContentLoaded', fn);
    }
  }

  ready(function () {
    var el = document.getElementById('chart-efficiency');
    if (!el || typeof echarts === 'undefined') {
      return;
    }

    // 与文档 CSS 变量保持一致的主题色
    var COLOR_BEFORE  = '#6B7280'; // 优化前 - 灰色
    var COLOR_AFTER   = '#065A82'; // 优化后 - 主题色
    var COLOR_INK     = '#1F2937';
    var COLOR_MUTED   = '#6B7280';
    var COLOR_ACCENT3 = '#E86A33';
    var COLOR_RULE    = '#E5E7EB';
    var FONT = "'Noto Sans SC', -apple-system, BlinkMacSystemFont, sans-serif";

    // 场景与数据（耗时统一换算为分钟）
    var scenarios = [
      '操作手册转培训PPT(20页)',
      '制度文档转宣贯PPT(15页)',
      '合规考点PPT生成(10页)',
      'PPT风格调整(整体换色)',
      '单页内容修正(补注意事项)'
    ];
    var beforeData = [1200, 720, 480, 240, 30]; // 优化前（分钟）
    var afterData  = [45, 30, 20, 15, 5];       // 优化后（分钟）
    var efficiency = ['96%', '94%', '96%', '94%', '83%']; // 提效幅度

    var chart = echarts.init(el, null, { renderer: 'canvas' });

    chart.setOption({
      backgroundColor: 'transparent',
      textStyle: { fontFamily: FONT, color: COLOR_INK },
      grid: { left: 48, right: 24, top: 44, bottom: 92, containLabel: true },
      legend: {
        data: ['优化前', '优化后'],
        top: 6,
        right: 8,
        itemWidth: 14,
        itemHeight: 14,
        itemGap: 18,
        textStyle: { fontFamily: FONT, color: COLOR_INK, fontSize: 12 }
      },
      tooltip: {
        trigger: 'axis',
        axisPointer: { type: 'shadow' },
        backgroundColor: '#FFFFFF',
        borderColor: COLOR_RULE,
        borderWidth: 1,
        padding: [8, 12],
        textStyle: { fontFamily: FONT, color: COLOR_INK, fontSize: 12 },
        formatter: function (params) {
          var idx = params[0].dataIndex;
          return params[0].name + '<br/>'
            + params[0].marker + ' 优化前：' + params[0].value + ' 分钟<br/>'
            + params[1].marker + ' 优化后：' + params[1].value + ' 分钟<br/>'
            + '<strong style="color:' + COLOR_ACCENT3 + '">提效：' + efficiency[idx] + '</strong>';
        }
      },
      xAxis: {
        type: 'category',
        data: scenarios,
        axisLine: { lineStyle: { color: '#9CA3AF' } },
        axisTick: { show: false },
        axisLabel: {
          fontFamily: FONT,
          color: COLOR_INK,
          fontSize: 11,
          interval: 0,
          rotate: 20,
          margin: 12
        }
      },
      yAxis: {
        type: 'value',
        name: '分钟',
        min: 0,
        max: 1500,
        interval: 300,
        nameTextStyle: { fontFamily: FONT, color: COLOR_MUTED, fontSize: 11, padding: [0, 0, 4, -28] },
        axisLine: { show: false },
        axisTick: { show: false },
        splitLine: { lineStyle: { color: COLOR_RULE, type: 'dashed' } },
        axisLabel: { fontFamily: FONT, color: COLOR_MUTED, fontSize: 11 }
      },
      series: [
        {
          name: '优化前',
          type: 'bar',
          data: beforeData,
          barWidth: 24,
          barGap: '25%',
          itemStyle: { color: COLOR_BEFORE, borderRadius: [4, 4, 0, 0] },
          // 在柱状图上方显示提效百分比标签
          label: {
            show: true,
            position: 'top',
            distance: 6,
            fontFamily: FONT,
            fontSize: 12,
            fontWeight: 600,
            color: COLOR_AFTER,
            formatter: function (params) {
              return '提效 ' + efficiency[params.dataIndex];
            }
          }
        },
        {
          name: '优化后',
          type: 'bar',
          data: afterData,
          barWidth: 24,
          itemStyle: { color: COLOR_AFTER, borderRadius: [4, 4, 0, 0] },
          // 优化后耗时数值标签（小柱也保持可读）
          label: {
            show: true,
            position: 'top',
            distance: 4,
            fontFamily: FONT,
            fontSize: 11,
            color: COLOR_AFTER,
            formatter: '{c} 分钟'
          }
        }
      ]
    });

    // 响应式：窗口尺寸变化时重绘
    window.addEventListener('resize', function () {
      chart.resize();
    });
  });
})();
