#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""重写 CSV 为正确的15列格式"""
import csv
import os

HEADER = ['type','key','value','desc','seq','node_type','content','shape',
          'width_cm','height_cm','bg_color','text_color','branch_to','branch_label','branch_kind']

CONFIGS = [
    ['config','title','个人综合签约业务流程','流程图标题'],
    ['config','preset','green','配色预设'],
    ['config','no_connectors','true','禁用连接线'],
    ['config','step_gap_cm','1.2','纵向间隔'],
    ['config','box_width_cm','5.0','矩形宽'],
    ['config','box_height_cm','0.6','矩形高'],
    ['config','diamond_width_cm','4.5','菱形宽'],
    ['config','diamond_height_cm','1.0','菱形高'],
    ['config','title_bg','1F3864','标题背景'],
    ['config','title_text','FFFFFF','标题文字色'],
]

NODES = [
    ['','1','main','客户到达网点取号','rect','5.0','0.6','C6EFCE','006100'],
    ['','2','main','Pad身份识别与分流','rect','5.0','0.6','C6EFCE','006100'],
    ['','3','main','柜员登录进入签约场景','rect','5.0','0.6','C6EFCE','006100'],
    ['','4','main','选择证件类型并读取','diamond','4.5','1.0','FFF2CC','7F6000','41','信息缺失','error'],
    ['','5','main','人脸识别身份核实','diamond','4.5','1.0','FFF2CC','7F6000','42','不通过','error'],
    ['','6','main','选择介质读取方式','rect','5.0','0.6','C6EFCE','006100'],
    ['','7','main','读取介质并校验状态','diamond','4.5','1.0','FFF2CC','7F6000','43','异常','error'],
    ['','8','main','介质验证输入密码','rect','5.0','0.6','C6EFCE','006100'],
    ['','9','main','服务类型选择与签约','rect','5.0','0.6','C6EFCE','006100'],
    ['','10','main','费用收取与短信通知','rect','5.0','0.6','C6EFCE','006100'],
    ['','11','main','交易确认与电子签名','diamond','4.5','1.0','FFF2CC','7F6000','44','签字不符','error'],
    ['','12','main','凭证打印与回单服务','rect','5.0','0.6','C6EFCE','006100'],
    ['','13','main','客户评价','rect','5.0','0.6','C6EFCE','006100'],
    ['','14','main','系统质检与自动归档','rect','5.0','0.6','C6EFCE','006100'],
    ['','41','branch','客户信息维护','rect','5.0','0.6','FCE4EC','C00000'],
    ['','42','branch','上级现场审核','rect','5.0','0.6','FCE4EC','C00000'],
    ['','43','branch','拒绝交易','rect','5.0','0.6','FCE4EC','C00000'],
    ['','44','branch','重新签字确认','rect','5.0','0.6','FCE4EC','C00000'],
]

def pad(row, n=15):
    return row + [''] * (n - len(row))

out_path = os.path.join(os.path.dirname(__file__), '..', '生成产物', '个人综合签约_节点表.csv')
with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(HEADER)
    for cfg in CONFIGS:
        writer.writerow(pad(cfg))
    for node in NODES:
        writer.writerow(pad(node))

print(f'Written: {out_path}')

# Verify
with open(out_path, encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    rows = list(reader)
    valid = [r for r in rows if r.get('seq') and r['seq'].strip()]
    print(f'Verify: {len(valid)} node rows parsed correctly')
    if valid:
        print(f'  First node: seq={valid[0]["seq"]}, content={valid[0]["content"]}')
