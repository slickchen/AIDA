# -*- coding: utf-8 -*-
"""
模板管理模块
"""

import json
import os


class TemplateManager:
    """模板管理器"""
    
    def __init__(self, config_dir='config'):
        self.config_dir = config_dir
        self.templates = self._load_templates()
    
    def _load_templates(self):
        """加载模板"""
        templates = {
            'standard': self._create_standard_template(),
            'minimal': self._create_minimal_template(),
            'hardware_only': self._create_hardware_template(),
            'software_only': self._create_software_template()
        }
        
        # 尝试从配置文件加载自定义模板
        config_file = os.path.join(self.config_dir, 'templates.json')
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    custom_templates = json.load(f)
                    templates.update(custom_templates)
            except:
                pass
        
        return templates
    
    def _create_standard_template(self):
        """创建标准模板（1:1还原您的排版）"""
        return {
            'name': '标准模板',
            'description': '完整还原AIDA64报告的所有数据',
            'sections': [
                {
                    'name': '系统概述',
                    'items': [
                        '计算机类型', '操作系统', '计算机名称', '用户名称', '登录域',
                        '处理器名称', '主板名称', '主板芯片组', '系统内存',
                        '显示适配器', '3D 加速器', '显示器'
                    ]
                },
                {
                    'name': '存储设备',
                    'items': [
                        '存储控制器1', '存储控制器2', '硬盘驱动器1', '硬盘驱动器2',
                        '硬盘 SMART 状态'
                    ]
                },
                {
                    'name': '磁盘分区',
                    'items': [
                        'C: (NTFS)', 'D: (NTFS)', 'E: (NTFS)', '总大小'
                    ]
                },
                {
                    'name': '网络设备',
                    'items': [
                        '主 IP 地址', '主 MAC 地址', '网络适配器1', '网络适配器2', '网络适配器3'
                    ]
                },
                {
                    'name': 'DMI信息',
                    'items': [
                        'DMI BIOS 厂商', 'DMI BIOS 版本', 'DMI 系统制造商', 'DMI 系统产品',
                        'DMI 系统版本', 'DMI 系统序列号', 'DMI 系统 UUID', 'DMI 主板制造商',
                        'DMI 主板产品', 'DMI 主板版本', 'DMI 主板序列号', 'DMI 主机制造商',
                        'DMI 主机版本', 'DMI 主机序列号', 'DMI 主机识别标签', 'DMI 主机类型'
                    ]
                },
                {
                    'name': '内存信息',
                    'items': [
                        'DIMM1: 模块名称', 'DIMM1: 序列号', 'DIMM1: 制造日期', 'DIMM1: 模块容量',
                        'DIMM1: 模块类型', 'DIMM1: 存取类型', 'DIMM1: 存取速度', 'DIMM1: 模块位宽',
                        'DIMM1: 模块电压', 'DIMM1: 错误检测方式', 'DIMM1: DRAM 制造商',
                        'DIMM3: 模块名称', 'DIMM3: 序列号', 'DIMM3: 制造日期', 'DIMM3: 模块容量',
                        'DIMM3: 模块类型', 'DIMM3: 存取类型', 'DIMM3: 存取速度', 'DIMM3: 模块位宽',
                        'DIMM3: 模块电压', 'DIMM3: 错误检测方式', 'DIMM3: DRAM 制造商'
                    ]
                },
                {
                    'name': '显示器信息',
                    'items': [
                        '显示器名称', '显示器 ID', '显示器型号', '显示器类型', '制造日期',
                        '序列号', '最大分辨率', '接口类型'
                    ]
                },
                {
                    'name': '网络详情',
                    'items': [
                        '网络适配器', 'MAC地址 (有线)', 'IP地址/子网掩码 (有线)',
                        '连接速度 (有线)', '已接收字节 (有线)', '已发送字节 (有线)'
                    ]
                },
                {
                    'name': '已安装程序',
                    'include': True,
                    'all_items': True
                }
            ]
        }
    
    def _create_minimal_template(self):
        """创建精简模板"""
        return {
            'name': '精简模板',
            'description': '仅包含关键系统信息',
            'sections': [
                {
                    'name': '核心硬件',
                    'items': [
                        '处理器名称', '主板名称', '系统内存', '显示适配器'
                    ]
                },
                {
                    'name': '存储',
                    'items': [
                        '硬盘驱动器1', '硬盘驱动器2', 'C: (NTFS)', 'D: (NTFS)', 'E: (NTFS)'
                    ]
                },
                {
                    'name': '网络',
                    'items': [
                        '主 IP 地址', '网络适配器1'
                    ]
                }
            ]
        }
    
    def _create_hardware_template(self):
        """创建硬件专用模板"""
        return {
            'name': '硬件专用模板',
            'description': '仅提取硬件信息',
            'sections': [
                {
                    'name': 'CPU和主板',
                    'items': [
                        '处理器名称', '主板名称', '主板芯片组'
                    ]
                },
                {
                    'name': '内存',
                    'items': [
                        '系统内存', 'DIMM1: 模块名称', 'DIMM1: 模块容量', 'DIMM1: 存取速度',
                        'DIMM3: 模块名称', 'DIMM3: 模块容量', 'DIMM3: 存取速度'
                    ]
                },
                {
                    'name': '显卡和显示器',
                    'items': [
                        '显示适配器', '3D 加速器', '显示器', '最大分辨率'
                    ]
                },
                {
                    'name': '存储',
                    'items': [
                        '硬盘驱动器1', '硬盘驱动器2', '硬盘 SMART 状态'
                    ]
                }
            ]
        }
    
    def _create_software_template(self):
        """创建软件专用模板"""
        return {
            'name': '软件专用模板',
            'description': '仅提取软件信息',
            'sections': [
                {
                    'name': '操作系统',
                    'items': [
                        '操作系统', '计算机名称', '用户名称'
                    ]
                },
                {
                    'name': '已安装程序',
                    'include': True,
                    'all_items': True
                }
            ]
        }
    
    def get_template(self, template_name):
        """获取模板"""
        return self.templates.get(template_name, self.templates['standard'])
    
    def get_available_templates(self):
        """获取可用模板列表"""
        return list(self.templates.keys())
    
    def get_template_items(self, template_name):
        """获取模板中的所有项目"""
        template = self.get_template(template_name)
        items = []
        
        for section in template.get('sections', []):
            if section.get('include', False) and section.get('all_items', False):
                # 包含所有项目
                items.append('已安装程序')
            else:
                items.extend(section.get('items', []))
        
        return items
    
    def save_custom_template(self, template_name, template_data):
        """保存自定义模板"""
        config_file = os.path.join(self.config_dir, 'templates.json')
        
        # 加载现有模板
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                templates = json.load(f)
        else:
            templates = {}
        
        # 添加新模板
        templates[template_name] = template_data
        
        # 保存
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(templates, f, ensure_ascii=False, indent=2)
        
        # 更新内存中的模板
        self.templates[template_name] = template_data
